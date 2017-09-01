var monthNames = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

function onOpen() {
  var ui = DocumentApp.getUi();
  // Or FormApp or SpreadsheetApp.
  ui.createMenu('Macro')
      .addItem('Update last modified date', 'replaceFormattedDate')
      .addToUi();
}

function replaceDate(regex, replacement) { // replace all
  var body = DocumentApp.getActiveDocument().getBody();
  return body.replaceText(regex, replacement);
}

function replaceFormattedDate() {
  var regex = "[a-zA-z]+\\s\\d{1,2},\\s\\d{4}", // escape '\', and matches 'monthName dd, yyyy' format
      formattedDate = DocumentApp.getActiveDocument().getBody().findText(regex),
      d,
      dd,
      monthName,
      yyyy,
      date,
      element;
  if (formattedDate) {
      d = new Date();
      dd = d.getDate();
      dd = pad(dd, 2)
      monthName = monthNames[d.getMonth()]; // Months are zero based, same with array
      yyyy = d.getFullYear();
      date = monthName + ' ' + dd + ', ' + yyyy;
      replaceDate(regex, date);
    } else {
      DocumentApp.getUi().alert('Cannot find a date format that matches "[a-zA-z]+\\s\\d{1,2},\\s\\d{4}" in the document.');
  }
}

function pad(str, max) {
  str = str.toString();
  return str.length < max ? pad("0" + str, max) : str;
}
