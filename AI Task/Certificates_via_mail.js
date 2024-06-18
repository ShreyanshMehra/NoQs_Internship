function sendCertificates() {

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var templateId = '1IkYvB9sWmKGyKiycPyJiZbbBeY2wZjsfsZuoLTJi1cQ'; 
    
    var templateFile = DriveApp.getFileById(templateId);
  
  
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var name = data[i][0];
      var email = data[i][2];
      var profile = data[i][3];
      var from_date = data[i][4];
      var to_date = data[i][6];
  
      var copyId = templateFile.makeCopy(name+' intern certificate').getId();
      var copyDoc = DocumentApp.openById(copyId);
      var body = copyDoc.getBody();
      
  
      body.replaceText('<<name>>', name);
      body.replaceText('<<profile>>', profile);
      body.replaceText('<<from date>>', from_date);
      body.replaceText('<<to date>>', to_date);
  
      copyDoc.saveAndClose();
      
      var pdf = DriveApp.getFileById(copyId);
      
  
      var subject = 'Dear ' + name + ': Enclosed herewtih the Certificate of your Internship';
      var message = 'Dear ' + name + ',\n\nThank you for completing internship at NoQs in the profile of '+ profile + '.\n\nWe wish you all the best in your future endeavours.\n\nThanking you,\nTeam HR';
      MailApp.sendEmail(email, subject, message, {
        attachments: [pdf]
      });
      
      DriveApp.getFileById(copyId).setTrashed(true);
    }
  }