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

function getOrCreateLabel(name: string) {
  var label = GmailApp.getUserLabelByName(name);

  if (!label) {
    label = GmailApp.createLabel(name);
  }

  return label;
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
  try {
    var todoLabel = getOrCreateLabel(Config.instance.unsubscribeLabel);
    var successLabel = getOrCreateLabel(Config.instance.successLabel);
    var failLabel = getOrCreateLabel(Config.instance.failLabel);

    const threads = todoLabel.getThreads();

    for (const thread of threads) {
      unsubscribeThread({
        thread,
        todoLabel,
        successLabel,
        failLabel,
      });
    }
  } catch (e) {
    Logger.log(String(e));
  }
}

function unsubscribeThread(args: {
  thread: GoogleAppsScript.Gmail.GmailThread;
  todoLabel: GoogleAppsScript.Gmail.GmailLabel;
  successLabel: GoogleAppsScript.Gmail.GmailLabel;
  failLabel: GoogleAppsScript.Gmail.GmailLabel;
}) {
  const { thread, todoLabel, successLabel, failLabel } = args;
  const message = thread.getMessages()[0];

  let status:
    | {
        summary: string;
        location: string;
      }
    | undefined = undefined;

  try {
    const raw = message.getRawContent();
    const rawMatch = raw.match(/^list\-unsubscribe:(.|\r\n\s)+<([^>]+)>/im);
    if (rawMatch) {
      const url = rawMatch[2];
      const protocol = url.match(/^([^:]+):/i)?.[1];
      const pathname = url.match(/^.+:([^?&#]+)/)?.[1];
      const subject = url.match(/[?&]subject=([^&]+)/i)?.[1] ?? "unsubscribe";
      const body = url.match(/[?&]body=([^&]+)/i)?.[1] ?? "unsubscribe";

      if (protocol?.startsWith("http")) {
        status = {
          summary: `Unsubscribed via header`,
          location: `POST to ${url}`,
        };
        UrlFetchApp.fetch(String(url), {
          method: "post",
        });
      }

      if (protocol === "mailto" && pathname) {
        const email = pathname;

        status = {
          summary: `Unsubscribed via email`,
          location: `${email} w/ subject "${subject.slice(0, 10)}..."`,
        };

        GmailApp.sendEmail(email, subject, body);
      }
    }

    if (!status) {
      const parseHrefRegex =
        /<a[^>]*href=["'](https?:\/\/[^"']+)["'][^>]*>(.*?)<\/a>/gi;
      const body = message.getBody().replace(/\s/g, "");
      let urls: RegExpExecArray | null = null;
      while ((urls = parseHrefRegex.exec(body))) {
        if (
          urls[1].match(/unsubscribe|optout|opt\-out|remove/i) ||
          urls[2].match(/unsubscribe|optout|opt\-out|remove/i)
        ) {
          const unsubscribeUrl = urls[1];
          status = {
            summary: "Maybe unsubscribed via link",
            location: unsubscribeUrl,
          };

          UrlFetchApp.fetch(unsubscribeUrl);
          break;
        }
      }
    }

    thread.removeLabel(todoLabel);
    if (status) {
      thread.addLabel(successLabel);
    } else {
      thread.addLabel(failLabel);
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

    const summary = status?.summary ? `Error in ${status.summary}` : "Error";
    const errorInfo =
      error instanceof Error ? error.stack ?? String(error) : String(error);
    const location = status?.location
      ? `in ${status.location}: ${errorInfo}`
      : errorInfo;

    logToSpreadsheet({
      from: message.getFrom(),
      subject: message.getSubject(),
      view: `=HYPERLINK("${thread.getPermalink()}", "View")`,
      unsubscribeLinkOrEmail: location,
      status: summary,
    });

    throw error;
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
          name: `Stop (running every ${CRON_MINUTES} minutes)`,
          functionName: "stopAllTriggers",
        }
      : {
          name: `Start running every ${CRON_MINUTES} minutes`,
          functionName: "startCronTrigger",
        },
    null,
    { name: "Settings...", functionName: "showConfigView" },
    {
      name: `  Unsubscribe threads labeled "${Config.instance.unsubscribeLabel}"`,
      functionName: "showConfigView",
    },
    {
      name: `  On success, label "${Config.instance.successLabel}"`,
      functionName: "showConfigView",
    },
    {
      name: `  On fail, label "${Config.instance.failLabel}"`,
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
}

type SerializedConfig = Readonly<Config & { runInBackground: boolean }>;

function getConfig(): SerializedConfig {
  return {
    unsubscribeLabel: Config.instance.unsubscribeLabel,
    successLabel: Config.instance.successLabel,
    failLabel: Config.instance.failLabel,
    runInBackground: isRunning(),
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
  return ScriptApp.getProjectTriggers().length > 0;
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

  updateMenu();
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

  updateMenu();
}
// ============================================================================
// Event Handlers
// ============================================================================

function onOpen() {
  showMenu();
}

function onInstall() {
  showMenu();
}
