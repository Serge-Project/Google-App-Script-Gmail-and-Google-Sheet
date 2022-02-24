var ui = SpreadsheetApp.getUi();
function myFunction() {
  ui.createMenu("Gmail Manager")
    .addItem("Get Emails by Label", "parse_email")
    .addToUi();
  ui.createMenu("Clear Cold")
    .addItem("Clear Cold", "Clear_cold_content")
    .addToUi();
  ui.createMenu("Clear Warm")
    .addItem("Clear Warm", "Clear_warm_content")
    .addToUi();
  ui.createMenu("Clear Mature")
    .addItem("Clear Mature", "Clear_Maturecontent")
    .addToUi();
  ui.createMenu("Clear All")
    .addItem("Clear All", "Clear_all_content")
    .addToUi();
}

function parse_email() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var cold = SS.getSheetByName("cold");
  var warm = SS.getSheetByName("warm");
  var mature = SS.getSheetByName("mature");
  var input = ui.prompt(
    "Label Name",
    "Enter the label name that is assigned to your emails:",
    Browser.Buttons.OK_CANCEL
  );
  var label = GmailApp.getUserLabelByName(input.getResponseText());
  var threads = label.getThreads();
  var mtr = false;

  if (input.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    if (messages.length == 1) {
      for (var j = 0; j < messages.length; j++) {
        var attachment = messages[j].getAttachments();
        var attachment_name = "";
        for (var k = 0; k < attachment.length; k++) {
          attachment_name = attachment_name + " - " + attachment[k].getName();
          var test_name = attachment[k].getName();
          Logger.log(test_name);
          if (mtr == false) {
            mtr = checkig_string(test_name);
          }
        }
        var date = messages[j].getDate();
        var subject = messages[j].getSubject();
        var sender = messages[j].getFrom();
        var receiver = messages[j].getTo();
        var content = messages[j].getPlainBody();
        var status = "Parent";

        if (mtr == true) {
          mature.appendRow([
            status,
            date,
            sender,
            receiver,
            subject,
            content,
            attachment_name,
          ]);
          mtr = false;
        } else {
          cold.appendRow([
            date,
            sender,
            receiver,
            subject,
            content,
            attachment_name,
          ]);
        }
      }
    } else {
      var status = "Parent";
      for (var j = 0; j < messages.length; j++) {
        var attachment = messages[j].getAttachments();
        var attachment_name = "";
        for (var k = 0; k < attachment.length; k++) {
          attachment_name = attachment_name + " - " + attachment[k].getName();
          var test_name = attachment[k].getName();
          Logger.log(test_name);
          if (mtr == false) {
            mtr = checkig_string(test_name);
          }
        }
        var date = messages[j].getDate();
        var subject = messages[j].getSubject();
        var sender = messages[j].getFrom();
        var receiver = messages[j].getTo();
        var content = messages[j].getPlainBody();

        if (mtr == true) {
          mature.appendRow([
            status,
            date,
            sender,
            receiver,
            subject,
            content,
            attachment_name,
          ]);
          mtr = false;
        } else {
          warm.appendRow([
            status,
            date,
            sender,
            receiver,
            subject,
            content,
            attachment_name,
          ]);
        }
        status = "Child";
      }
    }
    threads[i].removeLabel(label);
  }
}

function checkig_string(test_name) {
  var array = [
    ["noodoe ev distributor price list.pdf"],
    [
      "noodoe quotation",
      "quotation - noodoe",
      "quotation",
      "noodoe project quotation",
    ],
    ["invoice"],
    ["noodoe ev catalog"],
  ];

  var string_from_email_attachment = test_name.toLowerCase();

  for (var i = 0; i < 4; i++) {
    for (var j = 0; array[i][j] != null; j++) {
      var string_from_array = array[i][j];
      var length_of_string_from_array = string_from_array.length;
      for (var k = 0; k < length_of_string_from_array; k++) {
        if (string_from_array[k] != string_from_email_attachment[k]) {
          break;
        }
        if (k + 1 == length_of_string_from_array) {
          return true;
        }
      }
    }
  }
  return false;
}

function Clear_cold_content() {
  var range = SpreadsheetApp.getActive()
    .getSheetByName("Cold")
    .getRange("A2:G1000");
  range.clearContent();
}

function Clear_warm_content() {
  var range = SpreadsheetApp.getActive()
    .getSheetByName("Warm")
    .getRange("A2:G1000");
  range.clearContent();
}

function Clear_mature_content() {
  var range = SpreadsheetApp.getActive()
    .getSheetByName("Mature")
    .getRange("A2:G1000");
  range.clearContent();
}

function Clear_all_content() {
  var range = SpreadsheetApp.getActive()
    .getSheetByName("Cold")
    .getRange("A2:G1000");
  range.clearContent();

  var range = SpreadsheetApp.getActive()
    .getSheetByName("Warm")
    .getRange("A2:G1000");
  range.clearContent();

  var range = SpreadsheetApp.getActive()
    .getSheetByName("Mature")
    .getRange("A2:G1000");
  range.clearContent();
}
