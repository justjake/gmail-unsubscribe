/*
 * Gmail Unsubscribe
 * By Jake Teton-Landis (@jitl)
 *
 * Even if you're not a programmer, I hope the comments I've written are enough
 * to help you through everything that's going on.
 *
 * The important functions in this file are:
 *
 * - unsubscribeFromLabeledThreads: this is where we actually do the work
 *   of looking at emails and trying to unsubscribe from them.
 *
 * - onOpen: this function is called automatically when you open the Sheet.
 *   It just adds the menu to your sheet's UI.
 *   It's restricted from accessing your data by Google Apps Script itself.
 *
 * The best reference for automatic unsubscribing I've found is Google's
 * guide for mailers to set up one-click unsubscribe:
 * https://support.google.com/mail/answer/81126?hl=en#:~:text=Advanced%3A%20Set%20up%20one%2Dclick%20unsubscribe
 *
 * The relevant RFCs:
 * https://tools.ietf.org/html/rfc2369
 * https://tools.ietf.org/html/rfc8058
 *
 * @todo
 * I've seen Gmail offer to one-click unsubscribe from an email with only the
 * `list-unsubscribe-post` header and no `list-unsubscribe` header,
 * but I don't know where to send the POST request :(
 */

/**
 * @OnlyCurrentDoc
 */

// ============================================================================
// Gmail Unsubscriber
// ============================================================================

/**
 * @main
 * This is the function called from "Run Now", as well as called regularly if
 * you "start running periodically".
 */
function unsubscribeFromLabeledThreads() {
  var todoLabel = getOrCreateLabel(Config.instance.unsubscribeLabel);
  var successLabel = getOrCreateLabel(Config.instance.successLabel);
  var failLabel = getOrCreateLabel(Config.instance.failLabel);

  const threads = todoLabel.getThreads();

  for (const thread of threads) {
    try {
      unsubscribeThread({
        thread,
        todoLabel,
        successLabel,
        failLabel,
      });
    } catch (e) {
      console.log("Error in thread", {
        thread: thread.getPermalink(),
        error: e,
      });
    }
  }
}

/**
 * We parse these types of actions out of emails.
 * See the `switch` statement in `unsubscribeThread` for more details.
 */
type UnsubscribeAction =
  | { type: "http"; url: string; postBody?: string }
  | { type: "mailto"; email: string; subject?: string; emailBody?: string }
  | { type: "tryOpenLink"; url: string }
  | { type: "unknown"; url: string };

/**
 * We call this function on each thread tagged "Unsubscribe" (or whatever the
 * todo label is).
 */
function unsubscribeThread(args: {
  thread: GoogleAppsScript.Gmail.GmailThread;
  todoLabel: GoogleAppsScript.Gmail.GmailLabel;
  successLabel: GoogleAppsScript.Gmail.GmailLabel;
  failLabel: GoogleAppsScript.Gmail.GmailLabel;
}) {
  const { thread, todoLabel, successLabel, failLabel } = args;

  let status:
    | {
        summary: string;
        location: string;
        label?: "fail";
      }
    | undefined = undefined;

  let message: GoogleAppsScript.Gmail.GmailMessage | undefined = undefined;

  try {
    message = thread.getMessages()[0];

    /**
     * Parse the actions from the first message in the thread.
     * See the `getUnsubscribeActions` function for more details.
     */
    const actions = getUnsubscribeActions(message);
    console.log("Parsed thread", { thread: thread.getPermalink(), actions });

    /**
     * The first action is the best action (most likely to work).
     */
    const bestAction = actions.at(0);
    if (bestAction) {
      /**
       * Perform the action based on its `type`.
       */
      switch (bestAction.type) {
        /**
         * list-unsubscribe with an HTTP link: send POST to that URL as
         * appropriate. The request is sent from a Google-owned IP address.
         */
        case "http": {
          status = {
            summary: "Success via header",
            location: bestAction.postBody
              ? `POST to ${bestAction.url}\nwith body "${trimDetail(
                  bestAction.postBody
                )}"`
              : `GET ${bestAction.url}`,
          };

          UrlFetchApp.fetch(bestAction.url, {
            method: "post",
            payload: bestAction.postBody,
          });
          break;
        }

        /**
         * list-unsubscribe with a mailto link: send email to that address as
         * appropriate. The email is sent from your Gmail address.
         */
        case "mailto": {
          const parts = [bestAction.email];
          if (bestAction.subject) {
            parts.push(`w/ subject "${trimDetail(bestAction.subject)}"`);
          }
          if (bestAction.emailBody) {
            parts.push(`w/ body "${trimDetail(bestAction.emailBody)}"`);
          }
          status = {
            summary: "Success via email",
            location: parts.join("\n"),
          };

          GmailApp.sendEmail(
            bestAction.email,
            bestAction.subject ?? "unsubscribe",
            bestAction.emailBody ?? "unsubscribe"
          );
          break;
        }

        /**
         * <a href="url"> in the message body: send HTTP GET to that URL, in
         * hopes that that will unsubscribe us. The request is sent from a
         * Google IP address.
         *
         * We always mark this action as "fail" since we're uncertain if it
         * helps.
         */
        case "tryOpenLink": {
          status = {
            summary: "Maybe by opening link",
            location: bestAction.url,
            label: "fail",
          };

          UrlFetchApp.fetch(bestAction.url);
          break;
        }

        /**
         * We don't know how to handle this URL in the `list-unsubscribe` header.
         */
        case "unknown": {
          status = {
            summary: "Failed: don't know how",
            location: bestAction.url,
            label: "fail",
          };
          break;
        }

        /**
         * This should never happen.
         */
        default:
          throw new Error(
            `Parsed unknown action: ${JSON.stringify(bestAction)})`
          );
      }
    } else {
      status = {
        summary: "Failed: no action found",
        location: "",
        label: "fail",
      };
    }

    /**
     * Switch the label on the thread based on the status of the action.
     */
    thread.removeLabel(todoLabel);
    if (!status || status.label === "fail") {
      thread.addLabel(failLabel);
    } else {
      thread.addLabel(successLabel);
    }

    /**
     * Log to the spreadsheet.
     */
    logToSpreadsheet({
      from: message.getFrom(),
      subject: message.getSubject(),
      unsubscribeLinkOrEmail: status?.location || "not found",
      view: `=HYPERLINK("${thread.getPermalink()}", "View")`,
      status: status?.summary || "Could not unsubscribe",
    });
  } catch (error) {
    /**
     * On unexpected error, apply the "fail" label.
     */
    thread.addLabel(failLabel);
    thread.removeLabel(todoLabel);

    const summary = status?.summary
      ? `Error\nWhile "${status.summary.toLowerCase()}"`
      : "Error";
    const errorInfo =
      error instanceof Error ? error.stack ?? String(error) : String(error);
    const location = status?.location
      ? `in ${status.location}:\n${errorInfo}`
      : errorInfo;

    /**
     * Also log details of the error to the spreadsheet.
     */
    logToSpreadsheet({
      from: message?.getFrom() ?? "<error>",
      subject: message?.getSubject() ?? "<error>",
      view: `=HYPERLINK("${thread.getPermalink()}", "View")`,
      unsubscribeLinkOrEmail: location,
      status: summary,
    });

    throw error;
  }
}

/**
 * Parses a single Gmail message for the actions we could take to unsubscribe from it.
 * These actions are returned in order of preference.
 */
function getUnsubscribeActions(
  message: GoogleAppsScript.Gmail.GmailMessage
): UnsubscribeAction[] {
  const raw = message.getRawContent();
  /**
   * The best ways to unsubscribe are to follow the `list-unsubscribe` header
   * that the mailer adds to the email. This contains clear, machine-readable
   * instructions for how to unsubscribe.
   *
   * The header line may contain multiple unsubscribe actions, it looks like this:
   * ```
   * list-unsubscribe: <https://example.com/unsubscribe>, <mailto:unsubscribe@example.com?subject=XXXX>
   * ```
   * This regex grabs the contents of the header.
   */
  const listUnsubscribeHeader = raw.match(
    /**
     * This intimidating looking syntax is called a "regular expression"
     * (commonly shortened to "regex"). It's used for finding occurences of
     * specific patterns in a larger text.
     *
     * Try pasting into an online tool that will explain it letter by letter:
     * https://regexr.com/
     */
    /^list-unsubscribe:(?:[\r\n\s])+([^\n\r]+)$/im
  )?.[1];
  /**
   * If the unsubscribe action is a http/https URL, the mailer may include a
   * `list-unsubscribe-post` header, which specifies the body of the POST request
   * to send to the URL.
   */
  const listUnsubscribePostHeader = raw.match(
    /^list-unsubscribe-post:(?:[\r\n\s])+([^\n\r]+)$/im
  )?.[1];
  /**
   * Split the list-unsubscribe header contents into individual action URLs.
   *
   * input:
   * ```
   * "<https://example.com/unsubscribe>, <mailto:unsubscribe@example.com?subject=XXXX>"
   * ```
   * output:
   * ```
   * ["https://example.com/unsubscribe", "mailto:unsubscribe@example.com?subject=XXXX"]
   * ```
   */
  const listUnsubscribeOptionsMatches = listUnsubscribeHeader
    ? Array.from(listUnsubscribeHeader.matchAll(/<([^>]+)>/gi))
    : [];

  /**
   * Loop over the action URLs we found and parse out the details of each action.
   * We end up with an unsorted list of actions.
   */
  const actions: UnsubscribeAction[] = listUnsubscribeOptionsMatches.map(
    (match) => {
      const url = match[1];
      const protocol = url.match(/^([^:]+):/i)?.[1];
      const pathname = url.match(/^.+:([^?&#]+)/)?.[1];
      const subject = url.match(/[?&]subject=([^&]+)/i)?.[1];
      const emailBody = url.match(/[?&]body=([^&]+)/i)?.[1];
      const postBody = listUnsubscribePostHeader;

      if (protocol?.startsWith("http")) {
        return {
          type: "http",
          url: url,
          postBody,
        };
      }

      if (protocol === "mailto" && pathname) {
        const email = pathname;

        return {
          type: "mailto",
          email,
          subject,
          emailBody,
        };
      }

      return { type: "unknown", url };
    }
  );

  /**
   * This block of code attempts to find a clickable link in the HTML version of
   * the email, which is the version you usually see when you view an email in
   * Gmail.
   *
   * Any such link we find may unsubscribe you as soon as you visit, but it
   * could require you to fill out a form or click buttons to unsubscribe.
   *
   * We can "visit" these links by sending an HTTP GET request to them, in case
   * that's enough to automatically unsubscribe, but we can't be sure that's
   * effective.
   *
   * This is the lowest priority action.
   */
  const parseHrefRegex =
    /<a[^>]*href=["'](https?:\/\/[^"']+)["'][^>]*>(.*?)<\/a>/gi;
  const htmlBody = message.getBody().replace(/\s/g, "");
  let urls: RegExpExecArray | null = null;
  while ((urls = parseHrefRegex.exec(htmlBody))) {
    if (
      urls[1].match(/unsubscribe|optout|opt\-out|remove/i) ||
      urls[2].match(/unsubscribe|optout|opt\-out|remove/i)
    ) {
      actions.push({
        type: "tryOpenLink",
        url: urls[1],
      });
      break;
    }
  }

  return actions.sort((a, b) => getActionPriority(b) - getActionPriority(a));
}

/**
 * Rank the actions we find on an email. We'll perform the top ranked action found.
 *
 * "http" and "mailto" actions are parsed from the list-unsubscribe email header,
 * so the email sender is instructing us how to unsubscribe. These *should* work well.
 *
 * "tryOpenLink" just opens (HTTP GET) a "Unsubscribe" link probably meant for humans
 * in the HTML body, so it's the worst action.
 *
 * We also have an "unknown" action that just logs information. That and any new
 * actions we defined are ranked lowest.
 */
function getActionPriority(action: UnsubscribeAction): number {
  switch (action.type) {
    // This action is specified by the email sender as a way to
    // We rank HTTP post action highest, because it's the most efficient overall
    // if we can one-click unsubscribe via post.
    case "http":
      return action.postBody ? 3 : 1.5;
    // Sending an unsubscribe email is better than HTTP request if the
    // sender didn't indicate a one-click unsubscribe POST body.
    case "mailto":
      return 2;
    case "tryOpenLink":
      return 1;
    default:
      return 0;
  }
}

/**
 * Save the status of a thread to our spreadsheet.
 */
function logToSpreadsheet(args: {
  status: string;
  subject: string;
  view: string;
  from: string;
  unsubscribeLinkOrEmail: string;
}) {
  console.log("logToSpreadsheet:", args);
  var ss = SpreadsheetApp.getActive();
  ss.getActiveSheet().appendRow([
    args.status,
    args.subject,
    args.view,
    args.from,
    args.unsubscribeLinkOrEmail,
  ]);
}

/**
 * Called by the "Create labels" menu item.
 * Creates the Gmail labels used by this script.
 */
function createAllLabels() {
  getOrCreateLabel(Config.instance.unsubscribeLabel);
  getOrCreateLabel(Config.instance.successLabel);
  getOrCreateLabel(Config.instance.failLabel);
}

function trimDetail(detail: string) {
  if (detail.length > 40) {
    return String(detail).slice(0, 40) + "…";
  } else {
    return detail;
  }
}

function getOrCreateLabel(name: string) {
  var label = GmailApp.getUserLabelByName(name);

  if (!label) {
    label = GmailApp.createLabel(name);
  }

  return label;
}

// ============================================================================
// Menu
// ============================================================================

const MENU_NAME = "➡️ Gmail Unsubscriber";
function renderMenu() {
  return [
    { name: "Run Once", functionName: "unsubscribeFromLabeledThreads" },
    isRunning()
      ? {
          name: `Pause (running every ${CRON_MINUTES} minutes)`,
          functionName: "stopAllTriggers",
        }
      : {
          name: `Start running every ${CRON_MINUTES} minutes`,
          functionName: "startCronTrigger",
        },
    { name: "Create labels", functionName: "createAllLabels" },
    null,
    { name: "Settings...", functionName: "showConfigView" },
    {
      name: `- Unsubscribe threads labeled "${Config.instance.unsubscribeLabel}"`,
      functionName: "showConfigView",
    },
    {
      name: `- On success, label "${Config.instance.successLabel}"`,
      functionName: "showConfigView",
    },
    {
      name: `- On fail, label "${Config.instance.failLabel}"`,
      functionName: "showConfigView",
    },
  ];
}

function showMenu() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu(MENU_NAME, renderMenu());
}

function updateMenu() {
  SpreadsheetApp.getActiveSpreadsheet().updateMenu(
    MENU_NAME,
    renderMenu() as any
  );
}

// ============================================================================
// Config
// ============================================================================

class Config {
  static instance = new Config();

  private userProperties = PropertiesService.getUserProperties();

  get unsubscribeLabel(): string {
    return this.userProperties.getProperty("LABEL") || "Unsubscribe";
  }

  set unsubscribeLabel(value: string) {
    this.userProperties.setProperty("LABEL", value);
    updateMenu();
  }

  get successLabel(): string {
    return (
      this.userProperties.getProperty("SUCCESS_LABEL") || "Unsubscribe Success"
    );
  }

  set successLabel(value: string) {
    this.userProperties.setProperty("SUCCESS_LABEL", value);
    updateMenu();
  }

  get failLabel(): string {
    return (
      this.userProperties.getProperty("FAIL_LABEL") || "Unsubscribe Failed"
    );
  }

  set failLabel(value: string) {
    this.userProperties.setProperty("FAIL_LABEL", value);
    updateMenu();
  }

  get runInBackground(): boolean {
    return Boolean(
      this.userProperties.getProperty("RUN_IN_BACKGROUND") === "true"
    );
  }

  set runInBackground(value: boolean) {
    this.userProperties.setProperty("RUN_IN_BACKGROUND", String(value));
    updateMenu();
  }
}

type SerializedConfig = Readonly<Config>;

function getConfig(): SerializedConfig {
  return {
    unsubscribeLabel: Config.instance.unsubscribeLabel,
    successLabel: Config.instance.successLabel,
    failLabel: Config.instance.failLabel,
    runInBackground: Boolean(isRunning()),
  };
}

function saveConfig(config: SerializedConfig) {
  try {
    Config.instance.unsubscribeLabel = config.unsubscribeLabel;
    Config.instance.successLabel = config.successLabel;
    Config.instance.failLabel = config.failLabel;
    if (config.runInBackground) {
      startCronTrigger();
    } else {
      stopAllTriggers(false);
    }
    return `Updated config: ${JSON.stringify(getConfig(), null, 2)}`;
  } catch (e) {
    return "ERROR: " + e;
  }
}

function showConfigView() {
  var html = HtmlService.createHtmlOutputFromFile("config")
    .setTitle("Gmail Unsubscriber Settings")
    .setWidth(300)
    .setHeight(315)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

// ============================================================================
// Cron job
// ============================================================================
const CRON_MINUTES = 15;

function isRunning() {
  if (reducedPermissions) {
    // Can't access triggers during onOpen.
    return Config.instance.runInBackground;
  }

  try {
    return ScriptApp.getProjectTriggers().length > 0;
  } catch (error) {
    console.log("Cannot fetch triggers", error);
    return undefined;
  }
}

function startCronTrigger() {
  stopAllTriggers(true);

  ScriptApp.newTrigger("unsubscribeFromLabeledThreads")
    .timeBased()
    .everyMinutes(15)
    .create();

  Browser.msgBox(
    `Gmail Unsubscriber will run every ${CRON_MINUTES} minutes even if this spreadsheet is closed. You can stop it from the menu.`
  );

  Config.instance.runInBackground = true;
}

function stopAllTriggers(silent: boolean) {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  if (!silent) {
    Browser.msgBox(
      "The Gmail Unsubscriber has been disabled. You can restart it from the menu."
    );
  }

  Config.instance.runInBackground = false;
  updateMenu();
}

// ============================================================================
// Event Handlers
// ============================================================================

let reducedPermissions = false;
/**
 * This function runs automatically when you open the connected spreadsheet.
 * More info: https://developers.google.com/apps-script/guides/triggers#onopene
 */
function onOpen() {
  try {
    reducedPermissions = true;
    showMenu();
  } finally {
    reducedPermissions = false;
  }
}

function onInstall() {
  try {
    reducedPermissions = true;
    showMenu();
  } finally {
    reducedPermissions = false;
  }
}
