// Add menu item to Google Sheets to run the main function
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Update from Trello", functionName: "main"},];
  ss.addMenu("Trello", menuEntries);
}

// Trello variables
let enableLevelOfEffort = true; // Enable or disable level of effort
var api_key = "######"; // Grab your API key here https://trello.com/app-key
var api_token = "#####";
var board_id = "#####";
var url = "https://api.trello.com/1/";
var key_and_token = "key="+api_key+"&token="+api_token;

//Called by Google Sheet menu
function main() {
  var ss = SpreadsheetApp.getActiveSheet().clear();
  ss.appendRow(["Date", "Task", "Client", "Effort", "Who", "List", "Link"]); // List of column titles in row 1
  var response = UrlFetchApp.fetch(url + "boards/" + board_id + "/lists?cards=all&" + key_and_token); // Fetch the Trello feed
  var lists = JSON.parse((response.getContentText())); // Pass JSON into lists variable
  
  for (var i=0; i < lists.length; i++) {
    var list = lists[i];
      var response = UrlFetchApp.fetch(url + "list/" + list.id + "/cards?" + key_and_token);
      var cards = JSON.parse(response.getContentText());
      if(!cards) continue;
    for (var j=0; j < cards.length; j++) {
      var card = cards[j];
      var response = UrlFetchApp.fetch(url + "cards/" + card.id + "/?actions=all&" + key_and_token);
      var carddetails = JSON.parse(response.getContentText()).actions;
      if(!carddetails) continue;
      for (var k=0; k < carddetails.length; k++) {
        // Store required data into variables
        var dato = carddetails[k].date; // Card Date
        var fullname = carddetails[k].memberCreator.fullName; // Card creators full name
        var name = card.name; // Card title
        var link = card.url; // Link to card in Trello
        var listname = list.name; // Trello card name
        // Run loop to get the carld labels
        for (var q=0; q < card.labels.length; q++) {
          var labelname = (card.labels[q].name);
        }
        if (enableLevelOfEffort) {
          // Get Level of Effort from the Trello App (http://scrumfortrello.com/)
          var regExp = /\(([^)]+)\)/;
          var effort = regExp.exec(name);
          var loe;
          // Level of effort handling errors
          try {
            loe = effort[1];
          } catch (err) {
              loe = "N/A"
          } 
        }
      }
      ss.appendRow([dato, name, labelname, loe, fullname, listname, link]);
   }                                      
  }
}
