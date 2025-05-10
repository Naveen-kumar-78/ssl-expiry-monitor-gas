function updateSSLCertExpiry() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var today = new Date();
  var todayDateString = today.toDateString(); // Get only the date part
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  if (lastRow < 2) {
    Logger.log("No data found.");
    return;
  }

  var data = sheet.getRange(2, 2, lastRow - 1, 8).getValues(); // Columns B to I
  var expiryData = [];
  var projects = new Set();

  for (var i = 0; i < data.length; i++) {
    var rowIndex = i + 2; 
    var projectName = data[i][0]; // Column B (Project)
    var websiteUrl = data[i][1]; // Column C (Website URL)
    var expiryDate = data[i][5]; // Column G (SSL Expires On)

    if (typeof expiryDate === "string") {
      expiryDate = new Date(expiryDate.trim());
    }

    if (expiryDate && !isNaN(expiryDate.getTime())) {
      var daysUntilExpiry = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));
      var status, color;

      if (daysUntilExpiry <= 0) {
        status = "Expired";
        color = "grey";
      } else if (daysUntilExpiry <= 30) {
        status = "Expiring Soon";
        color = "red";
      } else if (daysUntilExpiry <= 50) {
        status = "Warning";
        color = "yellow";
      } else {
        status = "Active";
        color = "#ADD8E6";
      }

      sheet.getRange(rowIndex, 8).setValue(daysUntilExpiry); // Column H
      sheet.getRange(rowIndex, 9).setValue(status); // Column I
      sheet.getRange(rowIndex, 8, 1, 2).setBackground(color);

      if (daysUntilExpiry <= 30) {
        expiryData.push({
          project: projectName,
          url: websiteUrl,
          expiryDate: expiryDate.toDateString(),
          daysLeft: daysUntilExpiry
        });
        projects.add(projectName);
      }
    }
  }

  if (expiryData.length > 0) {
    sendExpiryAlert(expiryData, Array.from(projects), sheetUrl, todayDateString);
  }

  Logger.log("SSL Certificate Expiry Check Updated Successfully.");
}

function sendExpiryAlert(expiryData, projectNames, sheetUrl, todayDateString) {
  var recipients = ["naveen.kumar@bcits.in", "team_cloudl1support@bcits.co.in"];
  var projectList = projectNames.join(", ");
  var subject = `⚠️ SSL Expiry Alert – ${projectList} – Action Required`;

  var emailBody = `Dear Team,\n\nThe following SSL certificates are expiring soon:\n\n`;

  expiryData.forEach(function(entry) {
    emailBody += `🔹 **Project:** ${entry.project}\n`;
    emailBody += `🔹 **Website URL:** ${entry.url}\n`;
    emailBody += `🔹 **Expires In:** ${entry.daysLeft} days\n`;
    emailBody += `🔹 **Expiry Date:** ${entry.expiryDate}\n\n`;
  });

  emailBody += `📌 **Click here to view the full SSL certificate report:**\n➡️ ${sheetUrl}\n\n`;
  emailBody += `**Best Regards,**\nSSL Monitoring Team`;

  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSentDate = scriptProperties.getProperty("lastEmailSent");

  if (lastSentDate !== todayDateString) {
    MailApp.sendEmail({
      to: recipients.join(","),
      subject: subject,
      body: emailBody
    });

    scriptProperties.setProperty("lastEmailSent", todayDateString);
    Logger.log("Email sent successfully on " + todayDateString);
  } else {
    Logger.log("Email already sent today. Skipping.");
  }
}
