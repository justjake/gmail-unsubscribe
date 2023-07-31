/*
 * Gmail Unsubscribe
 *
 * By Jake Teton-Landis (@jitl)
 * Forked from Amit Agarwal (@labnol): https://www.labnol.org/internet/gmail-unsubscribe/28806/
 */

import "url-polyfill";

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
      const url = new URL(rawMatch[2]);

      if (url.protocol.startsWith("http")) {
        status = {
          summary: `Unsubscribed via header`,
          location: `POST to ${url}`,
        };
        UrlFetchApp.fetch(String(url), {
          method: "post",
        });
      }

      if (url.protocol === "mailto") {
        const email = url.pathname;
        const subject = url.searchParams.get("subject") || "Unsubscribe";
        const body = url.searchParams.get("body") || "Unsubscribe";

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
    { name: "Change labels...", functionName: "showConfigView" },
    {
      name: `  Unsubscribe labeled "${Config.instance.unsubscribeLabel}"`,
      functionName: "noOp",
    },
    {
      name: `  On success, label "${Config.instance.successLabel}"`,
      functionName: "noOp",
    },
    {
      name: `  On fail, label "${Config.instance.failLabel}"`,
      functionName: "noOp",
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

function getConfig(): Readonly<Config> {
  return {
    unsubscribeLabel: Config.instance.unsubscribeLabel,
    successLabel: Config.instance.successLabel,
    failLabel: Config.instance.failLabel,
  };
}

function saveConfig(config: Readonly<Config>) {
  try {
    Config.instance.unsubscribeLabel = config.unsubscribeLabel;
    Config.instance.successLabel = config.successLabel;
    Config.instance.failLabel = config.failLabel;
  } catch (e) {
    return "ERROR: " + e;
  }
}

function showConfigView() {
  var html = HtmlService.createHtmlOutputFromFile("config")
    .setTitle("Gmail Unsubscriber")
    .setWidth(300)
    .setHeight(500)
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
    "The Gmail Unsubscriber is now running in the background. You can stop it anytime later."
  );
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
}
// ============================================================================
// Event Handlers
// ============================================================================

function onOpen() {
  showMenu();
}
