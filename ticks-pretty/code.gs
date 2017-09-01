function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('My Menu')
    .addItem('Ticking', 'ticks')
    .addItem('triple ticking', 'tripleTicks')
    .addToUi();
  if (ScriptApp.getUserTriggers(DocumentApp.getActiveDocument())) {
    Logger.log('resetting triggers');
    ScriptApp.getUserTriggers(DocumentApp.getActiveDocument()).forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
    initiateProjTriggers();
  } else {
    initiateProjTriggers();
  }
}

function initiateProjTriggers() {
  ScriptApp.newTrigger('ticks')
    .forDocument(DocumentApp.getActiveDocument())
    .timeBased()
    .everyMinutes(1)
    .create();
  ScriptApp.newTrigger('tripleTicks')
    .forDocument(DocumentApp.getActiveDocument())
    .timeBased()
    .everyMinutes(1)
    .create();
}

function ticks() {
  var regex = "`{1}(\\w|\\d|[^`+])+`{1}", // double '\' to escape character
      ticksExist = DocumentApp.getActiveDocument().getBody().findText(regex),
      startPosition, endPosition;
  if (ticksExist) {
     if (ticksExist.isPartial()) {
       startPosition = ticksExist.getStartOffset();
       endPosition = ticksExist.getEndOffsetInclusive();
       ticksExist.getElement().asText().editAsText()
        .setFontFamily(startPosition, endPosition, 'Courier New')
        .setForegroundColor(startPosition, endPosition, '#FF0000')
        .setBackgroundColor(startPosition, endPosition, '#FFCDD2')
        .deleteText(startPosition, startPosition)
        .deleteText(endPosition - 1, endPosition - 1);
       ticks();
     } else {
       DocumentApp.getUi().alert('The entire range element is included.');
       Logger.log('The entire range element is included.');
     }
  }
}

function retrieveText(regex) {
  var body = DocumentApp.getActiveDocument().getBody();
  var foundElement = body.findText(regex);
  var set = [];
  while (foundElement != null) {
    // Aggregate
    set.push(foundElement);
    // Find the next match
    foundElement = body.findText(regex, foundElement);
  }
  return set;
}

function retrievetripleTicks(regex) {
  var ticksExist = DocumentApp.getActiveDocument().getBody().findText(regex),
      tickSet = retrieveText(regex);
  return tickSet;
}

function pairtripleTicks(tickSet) {
  var splitPairs = function(arr) {
        var pairs = [];
          for (var i=0 ; i<arr.length ; i+=2) {
            if (arr[i+1] !== undefined) {
              pairs.push ([arr[i], arr[i+1]]);
            } else {
              pairs.push ([arr[i]]);
            }
          }
          return pairs;
        };
  return splitPairs(tickSet);
}

function tripleTicks() {
  var regex = "`{3}",
      rangeBuilder = DocumentApp.getActiveDocument().newRange(),
      tickSet = retrievetripleTicks(regex),
      pairs = pairtripleTicks(tickSet),
      range,
      rangeElements;
  for (var i = 0, x = pairs.length; i < x; i++) {
    range = rangeBuilder.addElementsBetween(pairs[i][0].getElement(), pairs[i][1].getElement()).build();
    rangeElements = range.getRangeElements();
    rangeElements.forEach(function(rangeElement) {
      var element = rangeElement.getElement(),
          text = element.asText();
      text.editAsText()
        .setFontFamily('Courier New')
        .setForegroundColor('#01579B')
        .setBackgroundColor('#E1F5FE')
      text.replaceText(regex, '');
    })
  }
}
