// ====================================================================================================
// G4S KENNISTOETSEN - PDF GENERATOR & MAIL MODULE
// ====================================================================================================

function getLogoBase64() {
  try {
    // 1. Meest robuuste methode: Haal het bestand direct en veilig op uit Drive (omzeilt link restricties en browser blokkades)
    var file = DriveApp.getFileById(CONFIG.LOGO_ID);
    var blob = file.getBlob();
    var mime = blob.getContentType() || 'image/png';
    return "data:" + mime + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    Logger.log("Fout bij ophalen logo via DriveApp: " + e.message);
    // 2. Fallback methode
    try {
      var url = "https://drive.google.com/uc?export=download&id=" + CONFIG.LOGO_ID;
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var blob = response.getBlob();
      var mime = blob.getContentType() || 'image/png';
      // Check of google een inlogscherm stuurt (html) in plaats van de afbeelding
      if (mime.indexOf('text/html') === -1) {
          return "data:" + mime + ";base64," + Utilities.base64Encode(blob.getBytes());
      }
    } catch (err) {}
  }
  return "";
}

function createPDFBlob(d) {
  var color = d.finalStatus === 'Geslaagd' ? '#2E7D32' : '#D32F2F';
  var logoBase64 = getLogoBase64();
  
  // Lees de instelling uit die we vanuit Quiz.gs en JS.html hebben meegegeven (robuuste check)
  var toonAntwoorden = (d.toonAntwoorden !== false && String(d.toonAntwoorden).toUpperCase() !== "FALSE");

  var rows = "";
  if (d.details && d.details.length > 0) {
    for (var i = 0; i < d.details.length; i++) {
      var q = d.details[i];
      var antwoordInfo = q.givenAnswer;
      var resKleurBg = q.isCorrect ? '#F0FDF4' : '#FEF2F2';
      var resBorder = q.isCorrect ? '#bbf7d0' : '#fecaca';
      
      // Alleen het juiste antwoord tonen als het fout beantwoord was EN de vink aan staat
      var correctHtml = (q.isCorrect || !toonAntwoorden) ? '' : '<br><b style="color:#D32F2F;font-size:12px;">Correct: ' + q.correctAnswer + '</b>';
      var uitlegHtml = q.uitleg ? '<div style="margin-top: 10px; font-size: 13px; color: #475569; font-style: italic; background-color: #ffffff; padding: 8px; border-radius: 4px; border: 1px dashed #cbd5e1;"><b>Uitleg:</b> ' + q.uitleg + '</div>' : '';

      rows += '<div style="background-color: ' + resKleurBg + '; border: 1px solid ' + resBorder + '; padding: 15px; margin-bottom: 10px; border-radius: 6px; text-align: left; page-break-inside: avoid;">' +
        '<b style="color: #232323;">' + q.question + '</b><br>' +
        '<span style="font-size: 14px; color: #475569;">Jouw antwoord: ' + antwoordInfo + '</span>' +
        correctHtml +
        uitlegHtml +
      '</div>';
    }
  }

  var html = '<div style="font-family: Arial, sans-serif; padding: 20px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px; border-bottom: 2px solid #D32F2F; padding-bottom: 15px;">' +
      '<tr>' +
        '<td style="vertical-align: top;">' +
          '<h1 style="color:#232323; margin:0 0 5px 0;">RESULTAAT TOETS</h1>' +
          '<p style="margin:0;"><b>Kandidaat:</b> ' + d.userEmail + '<br><b>Toets:</b> ' + d.toetsNaam + '<br><b>Datum:</b> ' + (d.datum || '-') + '<br><b>Minimale score (norm):</b> ' + (d.norm || '-') + '%</p>' +
        '</td>' +
        '<td style="text-align: right; vertical-align: top;">' +
          '<img src="' + logoBase64 + '" style="height: 50px;">' +
        '</td>' +
      '</tr>' +
    '</table>' +
    '<div style="background:#f4f6f9; padding:25px; text-align:center; margin-bottom:20px; border-radius: 8px;">' +
      '<h1 style="font-size:50px; color:' + color + '; margin:0;">' + d.score + '%</h1>' +
      '<p style="color:' + color + '; font-weight:bold; font-size:18px; margin: 5px 0 0 0;">' + d.finalStatus + '</p>' +
    '</div>' +
    rows +
  '</div>';

  return HtmlService.createHtmlOutput(html).getAs('application/pdf').setName("Resultaat_" + d.userEmail + ".pdf");
}

function getPDFBase64(d) {
  var blob = createPDFBlob(d);
  return { base64: Utilities.base64Encode(blob.getBytes()) };
}

function sendPDFEmail(d) {
  if (!d || typeof d !== "object") {
    throw new Error("Geen rapportgegevens ontvangen.");
  }
  var recipientSource = String(d.sendToEmail || d.userEmail || "");
  var recipientMatch = recipientSource.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  if (!recipientMatch) {
    throw new Error("Geen geldig e-mailadres gevonden voor dit rapport.");
  }
  var to = recipientMatch[0];
  var subject = "Resultaat Toets: " + d.toetsNaam;
  // Robuuste check voor boolean en string
  var toonAntwoorden = (d.toonAntwoorden !== false && String(d.toonAntwoorden).toUpperCase() !== "FALSE");

  if (toonAntwoorden) {
    // Original behavior: Send email with PDF attachment
    var blob = createPDFBlob(d);
    var body = "Beste kandidaat,\n\nIn de bijlage vind je het resultaat en de foutenanalyse van jouw gemaakte toets.\n\nMet vriendelijke groet,\nG4S Operation Support Team";

    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body,
      attachments: [blob]
    });
  } else {
    // New behavior: Send simple text email without PDF
    var logoUrl = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : "https://www.g4s.com/-/media/g4s/global/images/logos/g4s---an-allied-universal-company-logo/g4s_one-company_logo.png";
    var color = d.finalStatus === 'Geslaagd' ? '#2E7D32' : '#D32F2F';

    var htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f6f9; padding: 20px; border-radius: 8px;">
        <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; border-top: 4px solid ${color}; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px;">
            <tr>
              <td style="vertical-align: middle;"><h2 style="color: #232323; font-size: 20px; margin: 0;">Resultaat Toets</h2></td>
              <td style="text-align: right; vertical-align: middle;"><img src="${logoUrl}" alt="G4S Logo" style="height: 50px;"></td>
            </tr>
          </table>
          <p style="color: #333; font-size: 15px; line-height: 1.6;">Beste kandidaat,</p>
          <p style="color: #333; font-size: 15px; line-height: 1.6;">Je hebt zojuist de toets <strong>${d.toetsNaam}</strong> afgerond.</p>
          
          <div style="background-color: #f8f9fa; padding: 20px; border-radius: 6px; margin: 20px 0; border: 1px solid #eee; text-align: center;">
            <h1 style="font-size: 40px; color: ${color}; margin: 0;">${d.score}%</h1>
            <p style="color: ${color}; font-weight: bold; font-size: 18px; margin: 10px 0 0 0; text-transform: uppercase;">${d.finalStatus}</p>
          </div>

          <div style="background-color: #e3f2fd; border-left: 4px solid #1565c0; padding: 15px; margin: 20px 0; border: 1px solid #cbe3f5; border-radius: 4px;">
             <p style="margin: 0; color: #0d47a1; font-size: 14px; line-height: 1.5;">Voor een gedetailleerde weergave van je resultaten kun je contact opnemen met je manager.</p>
          </div>
          
          <p style="color: #666; font-size: 13px; line-height: 1.5; margin-top: 30px;">
            Met vriendelijke groet,<br><strong>G4S Operation Support Team</strong>
          </p>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: htmlBody
    });
  }

  return { success: true };
}

function createGroupPDFBlob(payload) {
  var logoBase64 = getLogoBase64();
  var ges = payload.results.filter(function(r) { return r.status === 'Geslaagd'; }).length;
  var gez = payload.results.length - ges;

  var resultTableRows = "";
  if (payload.results && payload.results.length > 0) {
    resultTableRows = '<table width="100%" cellpadding="8" cellspacing="0" border="0" style="margin-bottom: 20px; font-size: 12px; border-collapse: collapse;">' +
      '<thead><tr style="background-color: #f4f6f9; border-bottom: 2px solid #e2e8f0; text-align: left; color: #475569;">' +
      '<th>Datum</th><th>Deelnemer</th><th>Score</th><th>Resultaat</th>' +
      '</tr></thead><tbody>';

    for (var j = 0; j < payload.results.length; j++) {
      var r = payload.results[j];
      var resKleur = r.status === 'Geslaagd' ? '#15803D' : '#D32F2F';
      var resDatum = r.displayDatum || r.datum || '-';
      resultTableRows += '<tr style="border-bottom: 1px solid #e2e8f0; page-break-inside: avoid;">' +
        '<td>' + resDatum + '</td>' +
        '<td>' + (r.email || '-') + '</td>' +
        '<td>' + (r.score || '-') + '%</td>' +
        '<td style="color: ' + resKleur + '; font-weight: bold;">' + (r.status || '-') + '</td>' +
        '</tr>';
    }
    resultTableRows += '</tbody></table>';
  } else {
    resultTableRows = '<p style="font-size: 13px; color: #666; font-style: italic;">Geen resultaten gevonden.</p>';
  }

  var statRows = "";
  if (payload.questionStats && payload.questionStats.length > 0) {
    for (var i = 0; i < payload.questionStats.length; i++) {
      var q = payload.questionStats[i];
      var percFout = Math.round((q.incorrect / q.total) * 100);
      
      statRows += '<div style="background-color: #FEF2F2; border: 1px solid #fecaca; padding: 15px; margin-bottom: 10px; border-radius: 6px; text-align: left; page-break-inside: avoid;">' +
        '<b style="color: #232323;">' + q.question + '</b><br>' +
        '<span style="font-size: 14px; color: #D32F2F; font-weight: bold;">' + q.incorrect + ' keer fout beantwoord (' + percFout + '% van totaal)</span>' +
      '</div>';
    }
  } else {
    statRows = '<div style="padding:15px; text-align:center; font-style:italic; color:#666;">Geen fouten gemaakt in deze selectie.</div>';
  }

  var html = '<div style="font-family: Arial, sans-serif; padding: 20px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px; border-bottom: 2px solid #D32F2F; padding-bottom: 15px;">' +
      '<tr>' +
        '<td style="vertical-align: top;">' +
          '<h1 style="color:#232323; margin:0 0 5px 0;">GROEPSRAPPORT TOETS</h1>' +
          '<p style="margin:0;"><b>Toets:</b> ' + payload.toets + '</p>' +
        '</td>' +
        '<td style="text-align: right; vertical-align: top;">' +
          '<img src="' + logoBase64 + '" style="height: 50px;">' +
        '</td>' +
      '</tr>' +
    '</table>' +
    '<div style="background:#f4f6f9; padding:25px; text-align:center; margin-bottom:20px; border-radius: 8px;">' +
      '<p style="font-size: 16px; margin:0;">Totaal: <b>' + payload.results.length + '</b> | Geslaagd: <b style="color:#2E7D32;">' + ges + '</b> | Gezakt: <b style="color:#D32F2F;">' + gez + '</b></p>' +
    '</div>' +
    '<h3 style="color:#232323; margin-top:10px; margin-bottom:15px; font-size:16px;">Resultaten Overzicht</h3>' +
    resultTableRows +
    '<h3 style="color:#232323; margin-top:30px; margin-bottom:15px; font-size:16px;">Foutenanalyse (Vaakst fout beantwoord)</h3>' +
    statRows +
  '</div>';

  return HtmlService.createHtmlOutput(html).getAs('application/pdf').setName("Groepsrapport_" + payload.toets + ".pdf");
}

function getGroupReportPDFBase64(payload) {
  var blob = createGroupPDFBlob(payload);
  return { base64: Utilities.base64Encode(blob.getBytes()) };
}

function sendGroupPDFEmail(payload, adminEmail) {
  var blob = createGroupPDFBlob(payload);
  var to = adminEmail;
  var subject = "Groepsrapport Toets: " + payload.toets;
  var body = "Beste beheerder,\n\nIn de bijlage vind je het gegenereerde groepsrapport en de foutenanalyse voor de geselecteerde toets.\n\nMet vriendelijke groet,\nG4S Operation Support Team";

  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: body,
    attachments: [blob]
  });
  return { success: true };
}