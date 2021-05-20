//&format=pdf                   //export format
//&size=a4                      //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
//&portrait=false               //true= Potrait / false= Landscape
//&scale=1                      //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
//&top_margin=0.00              //All four margins must be set!
//&bottom_margin=0.00           //All four margins must be set!
//&left_margin=0.00             //All four margins must be set!
//&right_margin=0.00            //All four margins must be set!
//&gridlines=false              //true/false
//&printnotes=false             //true/false
//&pageorder=2                  //1= Down, then over / 2= Over, then down
//&horizontal_alignment=CENTER  //LEFT/CENTER/RIGHT
//&vertical_alignment=TOP       //TOP/MIDDLE/BOTTOM
//&printtitle=false             //true/false
//&sheetnames=false             //true/false
//&fzr=false                    //true/false
//&fzc=false                    //true/false
//&attachment=false             //true/false

function test(){
  var deptSheets = ["Service","Maintenance","Drain","Sales","Install","Excavation"];
  for(sheet of deptSheets){
    createPDFReport(sheet,'MTD', '1Fc1aVyTnbwqWVoQqVrwXqkjcP89BF9e-', 'MTD' + ' - KPI Report - ' + sheet + ' - ')
  }
}


 function createPDFReport(sheetName, folderId, pdfName) {
  // https://stackoverflow.com/questions/26150732/creating-pdf-in-landscape-google-apps-script
  // https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  SpreadsheetApp.flush();
  
  var SHEETS = [sheetName];
  var FOLDERID = folderId
  var PDFNAME = pdfName
  
  var existingPDFs = DriveApp.getFilesByName(PDFNAME)
    while (existingPDFs.hasNext()) {
      var file = existingPDFs.next();
      file.setTrashed(true)
    }
  
  var FOLDER = DriveApp.getFolderById(FOLDERID);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    
  var sheet = ss.getSheetByName(sheetName);
  //Date set with format in EST (NYC) used in subject and PDF name
  var period = Utilities.formatDate(new Date(), "GMT+5", "YYYY-MM-dd");
  var url = ss.getUrl();
  
  //remove the trailing 'edit' from the url
  url = url.replace(/edit.*$/, '');

  //additional parameters for exporting the sheet as a pdf
  var url_ext = 'export?exportFormat=pdf&format=pdf' + //export as pdf
    //below parameters are optional...
    '&size=letter' + //paper size
    '&portrait=false' + //orientation, false for landscape
    '&fitw=true' + //fit to width, false for actual size
    '&sheetnames=false&printtitle=false&pagenumbers=false' + //hide optional headers and footers
    '&gridlines=false' + //hide gridlines
    '&fzr=true' + //do not repeat row headers (frozen rows) on each page
    '&top_margin=0.15&left_margin=0.25&right_margin=0.25&bottom_margin=0.25' + 
    '&pages=1' + 
    '&gid=' + sheet.getSheetId(); //the sheet's Id

  var token = ScriptApp.getOAuthToken();

  var response = UrlFetchApp.fetch(url + url_ext, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  var blob = response.getBlob().setName(PDFNAME);

  var newFile = FOLDER.createFile(blob);

  //from here you should be able to use and manipulate the blob to send and email or create a file per usual.
   var email = 'ecocommunications@ecoplumbers.com'; 
   var subject = "Client Concern Report " + period ;
   var body = "Please see the attached Client Concern report for last week.";
  //Place receipient email between the marks
  MailApp.sendEmail( email, subject, body, {attachments:[blob]});


}