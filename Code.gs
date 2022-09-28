let docIds = null;
let count = 0;
let lock = null;

function onMinuteInterval() {
  if (!obtainLock()) return;
  run();
  releaseLock();
}

function obtainLock() {
  lock = LockService.getScriptLock();
  if (!lock.tryLock(90000)) {
    Logger.log('Could not obtain script lock within 90 seconds');
    return false;
  }
  Logger.log('Script lock obtained');
  return true;
}

function releaseLock() {
  lock.releaseLock();
}

function run() {
  docIds = PropertiesService.getScriptProperties().getProperty('BoundDocIDs').split(',');
  let startTime = Date.now();
  while(Date.now() - startTime < 90000) {
    count++;
    processDocs();
  }
}

function processDocs() {
  docIds.forEach(docId => {
    try {
      Logger.log(count + ': Processing ' + docId);
      const doc = DocumentApp.openById(docId);
      processChildren(doc.getBody());
      doc.saveAndClose();
    } catch (e) {
      Logger.log(count + ': Could not bind to ' + docId);
    }
  });
}

function processChildren(parent) {
  if(!hasFunction(parent, 'getNumChildren')) return;
  const numChildren = parent.getNumChildren();
  for(var i = 0; i < numChildren; i++) {
    const child = parent.getChild(i);
    process(child);
    processChildren(child);
  }
}

function process(element) {
  const elementType = element.getType();
  switch(elementType) {
    case DocumentApp.ElementType.DATE:
      processDate(element); break;
  }
}

function processDate(date) {
  const previousSibling = date.getPreviousSibling();
  if (isReviewLabel(previousSibling)) {
    const timeSinceDate = timeSince(date.asDate().getTimestamp())
    const text = appendFriendlyText(date, timeSinceDate.friendlyText);
    formatFriendlyText(text, timeSinceDate);
  }
}

function appendFriendlyText(date, friendlyText) {
  const nextSibling = date.getNextSibling();

  if (isTextElement(nextSibling)) {
    return nextSibling.setText(' ' + friendlyText);
  }
  return date.getParent().asParagraph().appendText(' ' + friendlyText);
}

function formatFriendlyText(text, timeSinceDate) {
  if(timeSinceDate.intervalType === 'day') {
    if(timeSinceDate.interval < 0) text.setForegroundColor('#cc0000');
    else if(timeSinceDate.interval > 17) text.setForegroundColor('#bf9000');
    else text.setForegroundColor('#38761d');
  } else {
    text.setForegroundColor('#cc0000');
  }
}

function isReviewLabel(element) {
  return isTextElement(element)
    && element.asText().getText().trim() === 'Last reviewed:'
    && element.getParent().getType() === DocumentApp.ElementType.PARAGRAPH;
}

function isTextElement(element) {
  return hasFunction(element, 'asText');
}

function timeSince(date) {
  var seconds = Math.floor((new Date() - date) / 1000);
  var intervalType;

  var interval = Math.floor(seconds / 31536000);
  if (interval >= 1) {
    intervalType = 'year';
  } else {
    interval = Math.floor(seconds / 2592000);
    if (interval >= 1) {
      intervalType = 'month';
    } else {
      interval = Math.floor(seconds / 86400);
      intervalType = 'day';
    }
  }

  return {
    interval: interval,
    intervalType: intervalType,
    friendlyText: formatTimeSince(interval, intervalType)
  };
};

function formatTimeSince(interval, intervalType) {
  if (intervalType === 'day') {
    if (interval === 0) return 'today';
    if (interval === 1) return 'yesterday';
    if (interval < 0) return 'in the future';
  }

  if (interval > 1) intervalType += 's';
  return interval + ' ' + intervalType + ' ago'
}

function log(element) {
  const elementType = element.getType();
  Logger.log(elementType);
  switch(elementType) {
    case DocumentApp.ElementType.DATE:
      logDate(element); break;
    case DocumentApp.ElementType.RICH_LINK:
      logRichLink(element); break;
    case DocumentApp.ElementType.UNSUPPORTED:
      break;
    default:
      logText(element); break;
  }
}

function logDate(element) {
  Logger.log(element.asDate().getDisplayText());
}

function logRichLink(element) {
  Logger.log(element.asRichLink().getTitle());
}

function logText(element) {
  Logger.log(element.asText().getText());
}

function hasFunction(object, functionName) {
  return object !== null && typeof object[functionName] === 'function';
}