function sendBirthdayEmails() {

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    var data = sheet.getDataRange().getValues();
  
    for (var i = 1; i < data.length; i++) {
      var name = data[i][2];
      var email = data[i][4];
      var birthday_check = data[i][5];
  
      if (birthday_check == 'Yes') {
        try {
    
          var subject = 'Happy Birthday, ' + name + '!';
          var message = 'Dear ' + name + ',\n\nWishing you a very happy birthday!\n\nTeam HR ,\nNoQs Digital';
          MailApp.sendEmail(email, subject, message);
          
        } catch (e) {
          Logger.log('Error sending birthday email to ' + name + ': ' + e.message);
        }
      }
    }
  }
  
  function createDailyTrigger() {
    ScriptApp.newTrigger('sendBirthdayEmails')
             .timeBased()
             .everyDays(1)
             .atHour(8) 
             .create();
  }