/*
 * Gmail Unsubscribe
 *
 * By Jake Teton-Landis (@jitl)
 * Forked from Amit Agarwal (@labnol): https://www.labnol.org/internet/gmail-unsubscribe/28806/
 */

/**
 * @OnlyCurrentDoc
 */

// ============================================================================
// Gmail Unsubscriber
// ============================================================================

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

function createAllLabels() {
  getOrCreateLabel(Config.instance.unsubscribeLabel);
  getOrCreateLabel(Config.instance.successLabel);
  getOrCreateLabel(Config.instance.failLabel);
}

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

type UnsubscribeAction =
  | { type: "http"; url: string; postBody?: string }
  | { type: "mailto"; email: string; subject?: string; emailBody?: string }
  | { type: "tryOpenLink"; url: string }
  | { type: "unknown"; url: string };

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
    const actions = getUnsubscribeActions(message);
    console.log("Parsed thread", { thread: thread.getPermalink(), actions });

    const bestAction = actions.at(0);
    if (bestAction) {
      switch (bestAction.type) {
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

        case "tryOpenLink": {
          status = {
            summary: "Maybe by opening link",
            location: bestAction.url,
            label: "fail",
          };

          UrlFetchApp.fetch(bestAction.url);
          break;
        }

        case "unknown": {
          status = {
            summary: "Failed: don't know how",
            location: bestAction.url,
            label: "fail",
          };
          break;
        }

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

    thread.removeLabel(todoLabel);
    if (!status || status.label === "fail") {
      thread.addLabel(failLabel);
    } else {
      thread.addLabel(successLabel);
    }

    logToSpreadsheet({
      from: message.getFrom(),
      subject: message.getSubject(),
      unsubscribeLinkOrEmail: status?.location || "not found",
      view: `=HYPERLINK("${thread.getPermalink()}", "View")`,
      status: status?.summary || "Could not unsubscribe",
    });
  } catch (error) {
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

function getUnsubscribeActions(
  message: GoogleAppsScript.Gmail.GmailMessage
): UnsubscribeAction[] {
  const raw = message.getRawContent();
  const listUnsubscribeHeader = raw.match(
    /^list-unsubscribe:(?:[\r\n\s])+([^\n\r]+)$/im
  )?.[1];
  const listUnsubscribePostHeader = raw.match(
    /^list-unsubscribe-post:(?:[\r\n\s])+([^\n\r]+)$/im
  )?.[1];
  const listUnsubscribeOptionsMatches = listUnsubscribeHeader
    ? Array.from(listUnsubscribeHeader.matchAll(/<([^>]+)>/gi))
    : [];

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

function getActionPriority(action: UnsubscribeAction): number {
  switch (action.type) {
    case "http":
      return 3;
    case "mailto":
      return 2;
    case "tryOpenLink":
      return 1;
    default:
      return 0;
  }
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
    { name: "Create labels", functionName: "createAllLabels" },
  ];
}

function showMenu() {
  console.log("isOnOpen", isOnOpen);
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
  console.log("isRunning: isOnOpen", isOnOpen);
  if (isOnOpen) {
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
      "The Gmail Unsubscriber has been disabled. You can restart it anytime later."
    );
  }

  Config.instance.runInBackground = false;
  updateMenu();
}
// ============================================================================
// Event Handlers
// ============================================================================

let isOnOpen = false;
function onOpen() {
  try {
    isOnOpen = true;
    console.log("isOnOpen", isOnOpen);
    showMenu();
  } finally {
    isOnOpen = false;
  }
}

function onInstall() {
  showMenu();
}
