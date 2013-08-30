function main() {
  const count = 15;
  const filterTracking = 'from:ebay.com subject:Updates';
  const filter = 'from:ebay.com subject:Confirmation';

  var tracking = fetchLastGmailMessages_(filterTracking, count, processEbayTracking_);
  
  var items = fetchLastGmailMessages_(filter, count, processEbayConfirmation_);
  
  for(var i in items) {
    items[i].tracking = "";
    for each (track in tracking) {
      if (track.id == items[i].id) {
        items[i].tracking = track.tracking;
        break;
      }
    }
  }

  var spreadsheet = SpreadsheetApp.getActiveSheet();
  updateSpreadsheetWithItems_(spreadsheet, items);
}

function updateSpreadsheetWithItems_(sheet, items)
{
  for each (var item in items) {
    var row = [];
    for (var key in item) row.push(item[key]);
    sheet.appendRow(row);  
  }
}
 
function mergeItemListsById_(lhs, rhs) {
    for (var i in lhs) {
      
    lhs[i].tracking = "";
    for each (track in tracking) {
      if (track.id == items[i].id) {
        items[i].tracking = track.tracking;
        break;
      }
    }
  }
}

function fetchLastGmailMessages_(filter, count, processor)
{
  var threads = GmailApp.search(filter, 0, count);
  var items = [];
  Logger.log('Threads found: %s', threads.length);
  for (i in threads) {
    var messages = threads[i].getMessages();
    messages.forEach( function(message) {
      Logger.log(message.getDate());
      items.push(processor(message));
    } );
  }
  return items;
}

function processEbayTracking_(message) {
  var body = message.getBody();
  var item = new Object();

  var properties = [
    { name: "id", 
      query: { start: ">item: <a", end: "</a>", regex: /[\d]{10,}\t*$/ } },
    { name: "id", 
      query: { start: "\t( <a", end: "\t</a> )", regex: /[\d]{10,}\t*$/ } },
    { name: "tracking", 
      query: { start: "Tracking number:", end: "</a>", regex: /\w{8,}\s+$/ } }
  ];
  
  properties.forEach( function(property) {
    if (property.name in item) {
      Logger.log("Item already contains property '" + property.name + "': '" + item[property.name] + "'");
      return;
    }
    var value = performQuery_(body, property.query);
    if (value) {
      Logger.log("Add property '" + property.name + "': '" + value + "'");
      item[property.name] = value;
    }
    else Logger.log("Can't get property '"+property.name+"' for message dated " + message.getDate());
  });
  
  return item;
}

function processEbayConfirmation_(message) {
  var body = message.getBody();
  var item = new Object();
  
  var start = body.indexOf("Seller:");
  var end = body.indexOf("Email reference id:", start);
  if (start != -1 && end != -1)
  {
    body = body.slice(start, end);
  }
  else
  {
    Logger.log("Can't narrow message body");
  }
  
  var properties = [
    { name: "paid", 
      query: { start: "Paid on ", end: "</td>", regex: /\w+-\d{2}-\d{2}/ } },
    { name: "seller", 
      query: { start: "Seller:", end: "</a>", regex: /[\w\-\*\.]+$/ } },
    { name: "id", 
      query: { start: ">item: <a", end: "</a>", regex: /[\d]{10,}\t*$/ } },
    { name: "id", 
      query: { start: "\t( <a", end: "\t</a> )", regex: /[\d]{10,}\t*$/ } },
    { name: "title", 
      query: { start: "<b>Item total</b>", end: "</a>", regex: /[^>]+$/ } },
    { name: "price", 
      query: { start: "<td width=\"77\"", end: "</td>", regex: /\$\d+\.\d{2}/ } },
    { name: "shipping", 
      query: { start: "Shipping and handling", end: "</tr>", regex: /\$\d+\.\d{2}/, def: "$0" } },
    { name: "total", 
      query: { start: "Total", end: "</tr>", regex: /\$\d+\.\d{2}/ } }
  ];
  
  properties.forEach( function(property) {
    if (property.name in item) {
      return;
    }
    var value = performQuery_(body, property.query);
    if (value) {
      item[property.name] = value;
    }
    else Logger.log("Can't get property '"+property.name+"' for message dated " + message.getDate());
  });
  
  return item;
}

function performQuery_(text, query)
{
  var result = "def" in query? query.def : null;
  var start = text.indexOf(query.start);
  var end = text.indexOf(query.end, start);
  var cut = text.slice(start, end);
  if (cut) {
    var match = cut.match(query.regex);
    if (match)
      result = match[0].trim();
  }
  return result;
}

function cutItemBlock_(text) {
  const tail = 1024;
  var start = text.indexOf("Seller:");
  var end = text.indexOf("Tracking number:", start) + tail;
  return text.slice(start, end);
}
