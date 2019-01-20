function fill_sponsor_slide(){
  // Use the Sheets API to load data, one record per row.
  var dataRangeNotation = ''; //Put Data Range Here
  var dataSpreadsheetId = ''; //Put Google Sheet ID Here
  var templatePresentationId = ''; //Put Google Slide Template ID Here
  var values = SpreadsheetApp.openById(dataSpreadsheetId).getRange(dataRangeNotation).getValues();
  var spreadsheet = DriveApp.getFileById(dataSpreadsheetId);
  var j = 1; //Starting row for presentation link
  var sponsorname;
  var sponsorlogo;
  
  // For each record, create a new merged presentation.
  for (var i = 0; i < 42; i++) {
    sponsorname = values[i][0]; // name
    sponsorlogo = values[i][1]; //sponsor logo image link 
    var presentationcell = 'J' + j; //cell location for presentation link
    j++;
    
    // Duplicate the template presentation using the Drive API.
    var copyTitle = sponsorname + ' Slide Deck';
    var requests = {name:copyTitle
                   };
    var requests = {
      name: copyTitle
    };
    var driveResponse = Drive.Files.copy({
      resource: requests
    }, templatePresentationId);
    var presentationCopyId = driveResponse.id;
    var copy = DriveApp.getFileById(presentationCopyId);
    var setname = copy.setName(copyTitle);
    
    // Create the text merge (replaceAllText) requests for this presentation.
    requests = [{
      replaceAllText: {
        containsText: {
          text: '[INSERT COMPANY NAME]',
          matchCase: true
        },
        replaceText: sponsorname
      }
    },{
      replaceAllShapesWithImage: {
        imageUrl: sponsorlogo,
        replaceMethod: 'CENTER_INSIDE',
        containsText: {
          text: '[PARTNER LOGO]',
          matchCase: true
        }
      }
    }];              
    
    var url = copy.getUrl();
    var spread = SpreadsheetApp.openById(dataSpreadsheetId);
    var cell = spread.getRange(presentationcell);
    cell.setValue(url);
    copy.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    
    
    
    
    // Execute the requests for this presentation.
    var result = Slides.Presentations.batchUpdate({
      requests: requests
    }, presentationCopyId);
    
  }
}