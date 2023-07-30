/*

======================================================================

G M A I L   U N S U B S C R I B E

======================================================================

Last udpated on April 23, 2017

For help, send an email at amit@labnol.org

Tutorial: http://www.labnol.org/internet/amazon-price-tracker/28156/

*/

/**
 * @OnlyCurrentDoc
 */

function getConfig() {
  var params = {
    label: doProperty_("LABEL") || "Unsubscribe",
  };
  return params;
}

function config_() {
  var html = HtmlService.createHtmlOutputFromFile("config")
    .setTitle("Gmail Unsubscriber")
    .setWidth(300)
    .setHeight(200)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

function help_() {
  var html = HtmlService.createHtmlOutputFromFile("help")
    .setTitle("Google Scripts Support")
    .setWidth(350)
    .setHeight(120);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

function createLabel_(name) {
  var label = GmailApp.getUserLabelByName(name);

  if (!label) {
    label = GmailApp.createLabel(name);
  }

  return label;
}

function log_(status, subject, view, from, link) {
  var ss = SpreadsheetApp.getActive();
  ss.getActiveSheet().appendRow([status, subject, view, from, link]);
}

function init_() {
  Browser.msgBox(
    "The Unsubscriber was initialized. Please select the Start option from the Gmail menu to activate.",
  );
  return;
}

function onOpen() {
  var menu = [
    { name: "Configure", functionName: "config_" },
    null,
    { name: "☎ Help & Support", functionName: "help_" },
    { name: "✖ Stop (Uninstall)", functionName: "stop_" },
    null,
  ];

  SpreadsheetApp.getActiveSpreadsheet().addMenu("➪ Gmail Unsubscriber", menu);
}

function stop_(e) {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  if (!e) {
    Browser.msgBox(
      "The Gmail Unsubscriber has been disabled. You can restart it anytime later.",
    );
  }
}

function doProperty_(key, value) {
  var properties = PropertiesService.getUserProperties();

  if (value) {
    properties.setProperty(key, value);
  } else {
    return properties.getProperty(key);
  }
}

function doGmail() {
  try {
    var label = doProperty_("LABEL") || "Unsubscribe";

    var threads = GmailApp.search("label:" + label);

    var todoLabel = createLabel_(label);
    var successLabel = createLabel_("Unsubscribe Ok");
    var failLabel = createLabel_("Unsubscribe Fail");

    var url, urls, message, raw, body, formula, status;

    var hyperlink = '=HYPERLINK("#LINK#", "View")';

    var hrefs = new RegExp(
      /<a[^>]*href=["'](https?:\/\/[^"']+)["'][^>]*>(.*?)<\/a>/gi,
    );

    for (var t in threads) {
      url = "";

      status = "Could not unsubscribe";

      message = threads[t].getMessages()[0];

      threads[t].removeLabel(todoLabel);

      raw = message.getRawContent();

      urls = raw.match(/^list\-unsubscribe:(.|\r\n\s)+<(https?:\/\/[^>]+)>/im);

      if (urls) {
        url = urls[2];
        status = "Unsubscribed via header";
      } else {
        body = message.getBody().replace(/\s/g, "");
        while (url === "" && (urls = hrefs.exec(body))) {
          if (
            urls[1].match(/unsubscribe|optout|opt\-out|remove/i) ||
            urls[2].match(/unsubscribe|optout|opt\-out|remove/i)
          ) {
            url = urls[1];
            status = "Unsubscribed via link";
          }
        }
      }

      if (url === "") {
        urls = raw.match(/^list\-unsubscribe:(.|\r\n\s)+<mailto:([^>]+)>/im);
        if (urls) {
          url = parseEmail_(urls[2]);
          var subject = "Unsubscribe";
          GmailApp.sendEmail(url, subject, subject);
          status = "Unsubscribed via email";
        }
      }

      if (status.match(/unsubscribed/i)) {
        UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        threads[t].addLabel(successLabel);
      } else {
        threads[t].addLabel(failLabel);
      }

      formula = hyperlink.replace("#LINK", threads[t].getPermalink());

      log_(status, message.getSubject(), formula, message.getFrom(), url);
    }
  } catch (e) {
    Logger.log(e.toString());
  }
}

function saveConfig(params) {
  try {
    doProperty_("LABEL", params.label);

    stop_(true);

    ScriptApp.newTrigger("doGmail").timeBased().everyMinutes(15).create();

    return (
      "The Gmail unsubscriber is now active. You can apply the Gmail label " +
      params.label +
      " to any email and you'll be unsubscribed in 15 minutes. Please close this window."
    );
  } catch (e) {
    return "ERROR: " + e.toString();
  }
}

function parseEmail_(email) {
  var result = email.trim().split("?");
  return result[0];
}
