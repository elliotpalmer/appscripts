// EMAIL CODE

function emailATestPDFFile(){
  // Test out the Email Script
  sendEmail('elliot@ecoplumbers.com', 'Test','This is a test','KPI REPORT - MTD - 2019-10-1');
};

function sendEmail(to, subject, htmlBody, pdfFileName){
  // Simple Emailling of a PDF based on a PDF Name in your Drive
  var pdfBlob = getPDF(pdfFileName);
  
  MailApp.sendEmail({
  to: to,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [pdfBlob]
    });
  
  };


function getPDF(fileName){
  // Search the Drive
  var files = DriveApp.getFilesByName(fileName)
  
  // Get the PDF File (Blob)
  var pdfBlob = files.next().getBlob().getAs('application/pdf')
  
  return pdfBlob
  
  };
  
function testMultiPDFs(){
  var fileQuery = 'title contains "KPI" and title contains "2019-10-03.pdf"'
  
  sendEmailMultiAttach('elliot@ecoplumbers.com', 'Test','This is a test',fileQuery)
  
};

function sendEmailMultiAttach(to, subject, htmlBody, fileQuery){
  //  Adjusted Code for emailing multiple PDFs
  var pdfBlobs = getMultiplePDFs(fileQuery);
  
  MailApp.sendEmail({
    to: to,
    subject: subject,
    htmlBody: htmlBody,
    attachments: pdfBlobs
    });
  
  };


function getMultiplePDFs(fileQuery){
  
  // Sometimes a search query will return multiple results
  var files = DriveApp.searchFiles(fileQuery)
  var pdfBlobs = []
  
  // Loop through each file and add it to the list
  while (files.hasNext()) {
    var file = files.next();
    Logger.log(file.getName());
    
    var pdfBlob = file.getBlob().getAs('application/pdf')
    
    pdfBlobs.push(pdfBlob)
  };
  
  return(pdfBlobs)
  
};


