// run this code agest your Gmail account, it will keep list in new google spreadsheet in you Drive with file name "List of Emails in Gmail Account"
// please ask me if you have any question
function pullEmails() {
  var myEmail = Session.getActiveUser().getEmail();
  var createNewSpreadsheet = SpreadsheetApp.create("List of Emails in Gmail Account");
  var currentSheet = createNewSpreadsheet.getSheetByName("Sheet1").setName("AllGmailList");
  var startIndex = 0;
  var pageSize = 500;

  var listoemailAddresses = [];

  while (1) {
    var gmailThreads = GmailApp.search("from:(" + myEmail + ") || to:(" + myEmail + ")", startIndex, pageSize)

    // Get all messages for the current batch of threads
    var messages = GmailApp.getMessagesForThreads(gmailThreads);
    // Loop over all messages
    for (var i = 0; i < messages.length; i++) {
      // Loop over all messages in this thread
      for (var j = 0; j < messages[i].length; j++) {
        var mailFrom = messages[i][j].getFrom();
        var mailTo = messages[i][j].getTo()

        var matchesFrom = mailFrom.match(/\s*"?([^"]*)"?\s+<(.+)>/);
        var matchesTo = mailTo.match(/\s*"?([^"]*)"?\s+<(.+)>/);
        matchesFrom = matchesFrom && matchesFrom[2].toLocaleLowerCase()
        matchesTo = matchesTo && matchesTo[2].toLocaleLowerCase()

        if (matchesFrom && listoemailAddresses.indexOf(matchesFrom) == -1 && matchesFrom.indexOf("reply") == -1) {
          listoemailAddresses.push(matchesFrom)
        }
        if (matchesTo && listoemailAddresses.indexOf(matchesTo) == -1 && matchesTo.indexOf("reply") == -1) {
          listoemailAddresses.push(matchesTo)
        }
      }
    }

    if (gmailThreads.length == 0) {
      break
    } else {
      startIndex += pageSize;
    }
  }

  currentSheet.getRange("A:A").clear();
  listoemailAddresses = listoemailAddresses.map(val => [val]);
  currentSheet.getRange(1, 1, listoemailAddresses.length, listoemailAddresses[0].length).setValues(listoemailAddresses);

}
