// Polyfill
// https://tc39.github.io/ecma262/#sec-array.prototype.find
if (!Array.prototype.find) {
  Object.defineProperty(Array.prototype, "find", {
    value: function(predicate) {
      // 1. Let O be ? ToObject(this value).
      if (this == null) {
        throw new TypeError('"this" is null or not defined');
      }

      var o = Object(this);

      // 2. Let len be ? ToLength(? Get(O, "length")).
      var len = o.length >>> 0;

      // 3. If IsCallable(predicate) is false, throw a TypeError exception.
      if (typeof predicate !== "function") {
        throw new TypeError("predicate must be a function");
      }

      // 4. If thisArg was supplied, let T be thisArg; else let T be undefined.
      var thisArg = arguments[1];

      // 5. Let k be 0.
      var k = 0;

      // 6. Repeat, while k < len
      while (k < len) {
        // a. Let Pk be ! ToString(k).
        // b. Let kValue be ? Get(O, Pk).
        // c. Let testResult be ToBoolean(? Call(predicate, T, « kValue, k, O »)).
        // d. If testResult is true, return kValue.
        var kValue = o[k];
        if (predicate.call(thisArg, kValue, k, o)) {
          return kValue;
        }
        // e. Increase k by 1.
        k++;
      }

      // 7. Return undefined.
      return undefined;
    }
  });
}

// https://tc39.github.io/ecma262/#sec-array.prototype.findIndex
if (!Array.prototype.findIndex) {
  Object.defineProperty(Array.prototype, "findIndex", {
    value: function(predicate) {
      // 1. Let O be ? ToObject(this value).
      if (this == null) {
        throw new TypeError('"this" is null or not defined');
      }

      var o = Object(this);

      // 2. Let len be ? ToLength(? Get(O, "length")).
      var len = o.length >>> 0;

      // 3. If IsCallable(predicate) is false, throw a TypeError exception.
      if (typeof predicate !== "function") {
        throw new TypeError("predicate must be a function");
      }

      // 4. If thisArg was supplied, let T be thisArg; else let T be undefined.
      var thisArg = arguments[1];

      // 5. Let k be 0.
      var k = 0;

      // 6. Repeat, while k < len
      while (k < len) {
        // a. Let Pk be ! ToString(k).
        // b. Let kValue be ? Get(O, Pk).
        // c. Let testResult be ToBoolean(? Call(predicate, T, « kValue, k, O »)).
        // d. If testResult is true, return k.
        var kValue = o[k];
        if (predicate.call(thisArg, kValue, k, o)) {
          return k;
        }
        // e. Increase k by 1.
        k++;
      }

      // 7. Return -1.
      return -1;
    }
  });
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Blast email')
      .addItem('Send as ' + SENDER, 'sendEmails')
      .addItem('Review email template', 'reviewTemplate')
      .addToUi();
}

// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = new Date().toLocaleDateString(),
  SUBJECT = "Chat about your planning needs? - MagicBus",
  CC = "chris@magicbus.io",
  BCC = null,
  SENDER = toProperCase(Session.getEffectiveUser().getEmail().split("@")[0]),
  FOOTER = "MagicBus is a demand-responsive shuttle platform. Our system adapts to fixed and dynamic routing. As a proof of concept, we run commuter shuttles between cities and suburbs in the Bay Area and Detroit.\n\nVisit us at https://www.magicbus.io",
  MESSAGE_TEMPLATE = function(firstName, customMessage) {
    return "Hi " + firstName + ",\n\n" +
    "I saw you attended California Transportation Planning Conference back in May. My team has been working with corporate shuttle programs, and have started talking with transit agencies to learn more about the needs int he public space.\n\n" +
    customMessage + " Would you be open to spending 15-20 minutes on a call with me, to help us identify general trends and directions in the public space?\n\n" +
    "Please let me know. I’d love to set up a time to give you a call next week!\n\n" +
    "Best,\n" +
    SENDER + "\n\n" +
    FOOTER;
  }

function reviewTemplate() {
  SpreadsheetApp.getUi().alert(MESSAGE_TEMPLATE("FIRST_NAME", "CUSTOM_MESSAGE_GOES_HERE"))
}

function toProperCase(word) {
  var chars = word.split("");
  chars.splice(0, 1, chars[0].toUpperCase());
  return chars.join("");
}

function columnPosition(headerName) {
  var sheet = SpreadsheetApp.getActiveSheet(),
    values = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn()),
    columnPosition;
  columnPosition = values[0].findIndex(function(element) {
    return element === headerName;
  });
  return columnPosition;
}

function retrieveMessage(firstName, customMessage) {
  var message = MESSAGE_TEMPLATE(firstName, customMessage);
  return message;
}

function sendEmails() {
  var ui = SpreadsheetApp.getUi(),
    sheet = SpreadsheetApp.getActiveSheet(),
    startRow = 2, // First row of data to process
    numRows = sheet.getLastRow(), // all rows with content to process
    // Fetch values for each row in the Range.
    data = sheet.getSheetValues(startRow, 1, numRows, sheet.getLastColumn()),
    confirmed;

  confirmed = ui.alert("Are you sure you want to continue?", ui.ButtonSet.YES_NO)

  if (confirmed !== ui.Button.YES) {
    return;
  }

  for (var i = 0; i < data.length; ++i) {
    var row = data[i],
      firstName = row[columnPosition("First Name")],
      emailAddress = row[columnPosition("Email")],
      customMessage = row[columnPosition("Custom Message")],
      // TODO retrieve content from a function
      message = retrieveMessage(firstName, customMessage);
      emailSent = row[columnPosition("Sent Date")];
    if (!emailAddress) {
      return;
    }
    if (!emailSent) {
      // Prevents sending duplicates
      var subject = SUBJECT;
      MailApp.sendEmail(emailAddress, subject, message, {
        cc: CC,
        bcc: BCC,
        name: SENDER + " (MagicBus)"
      });
      sheet
        .getRange(startRow + i, columnPosition("Sent Date") + 1)
        .setValue(EMAIL_SENT); // columnPosition returns zero index
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
