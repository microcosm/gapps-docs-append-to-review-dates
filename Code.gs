const intervals = [
  { label: 'year', seconds: 31536000 },
  { label: 'month', seconds: 2592000 },
  { label: 'day', seconds: 86400 },
  { label: 'hour', seconds: 3600 },
  { label: 'minute', seconds: 60 },
  { label: 'second', seconds: 1 }
];

function onMinuteInterval() {
  processChildren(DocumentApp.getActiveDocument().getBody());
}

function processChildren(parent) {
  if(typeof parent.getNumChildren !== "function") return;
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

function processDate(element) {
  const timestamp = element.asDate().getTimestamp();
  Logger.log(timestamp + timeSince(timestamp));
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

  return formatTimeSince(interval, intervalType);
};

function formatTimeSince(interval, intervalType) {
  if (intervalType === 'day') {
    if (interval === 0) return ' is today';
    if (interval < 0) return ' is in the future';
  }

  if (interval > 1) intervalType += 's';
  return ' is ' + interval + ' ' + intervalType + ' ago'
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