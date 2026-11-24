// ====================================================================================================
// G4S KENNISTOETSEN - KANDIDAAT LOGICA (TOETS AFNAME)
// ====================================================================================================

function getQuizAttemptsSheet_(ss) {
  var sh = ss.getSheetByName(TAB_POGINGEN);
  if (!sh) {
    sh = ss.insertSheet(TAB_POGINGEN);
    sh.appendRow(["Token", "ToetsID", "SessieID", "Email", "VraagIDs", "Norm", "Starttijd", "MaxTijd", "Uitnodiging", "Status"]);
    sh.hideSheet();
  }
  return sh;
}

function getQuizData(toetsId, sid) {
  if (!toetsId) {
    return null;
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var email = getCurrentUserEmail_();
  if (!sid && !email) {
    throw new Error("Je Google-account kon niet worden vastgesteld. Meld je opnieuw aan en probeer het nogmaals.");
  }
  var reqCount = 0;
  var reqNorm = 70;
  var maxTijd = 0; // Toegevoegd voor tijdslimiet
  var startedAt = new Date();
  
  if (sid) {
    var shSes = ss.getSheetByName(TAB_SESSIES);
    
    if (!shSes) {
      throw new Error("Database sessies niet gevonden.");
    }
    
    var sRows = shSes.getDataRange().getValues();
    var valid = false;
    
    for (var i = sRows.length - 1; i >= 1; i--) {
      if (String(sRows[i][9]) === sid && sRows[i][1] == toetsId) {
        if (sRows[i][10] === "USED") {
          throw new Error("Deze link is al gebruikt. Je hebt de toets al gemaakt.");
        }
        
        valid = true; 
        reqNorm = parseNorm(sRows[i][2]);
        reqCount = parseInt(sRows[i][12], 10) || 0; 
        maxTijd = parseFloat(String(sRows[i][6]).replace(',', '.')) || 0; // Haal de tijd op uit kolom G (index 6)
        
        if (sRows[i][10] === "BEZIG" && sRows[i][11] instanceof Date) {
          startedAt = sRows[i][11];
        } else {
          shSes.getRange(i+1, 11).setValue("BEZIG");
          shSes.getRange(i+1, 12).setValue(startedAt);
        }
        
        break;
      }
    }
    
    if (!valid) {
      throw new Error("Ongeldige of verlopen link.");
    }
  } else {
    var shRes = ss.getSheetByName(TAB_RESULTATEN);
    
    if(shRes) {
      var lastRow = shRes.getLastRow();
      if (lastRow > 1) {
        // OPTIMALISATIE: We laden alleen de kolommen in die we echt nodig hebben (A t/m H)
        // Dit bespaart een enorme hoeveelheid geheugen omdat de JSON data in kolom I wordt overgeslagen
        var rRows = shRes.getRange(2, 1, lastRow - 1, 8).getValues();
        for (var j = 0; j < rRows.length; j++) {
          if (String(rRows[j][0]).toLowerCase() === email && rRows[j][7] == toetsId) {
            throw new Error("Je hebt deze toets al gemaakt.");
          }
        }
      }
    }
  }

  var shIdx = ss.getSheetByName(TAB_INDEX);
  
  if(!shIdx) {
    throw new Error("Database index niet gevonden.");
  }
  
  var idx = shIdx.getDataRange().getValues();
  var meta = null;
  
  for(var k = 1; k < idx.length; k++) {
    if(idx[k][0] == toetsId) { 
      meta = {
        id: idx[k][0], 
        naam: idx[k][1], 
        norm: (sid ? reqNorm : parseNorm(idx[k][7])), 
        aantal: reqCount || idx[k][8],
        maxTijd: maxTijd, // Stuur de maxTijd door naar het startscherm (JS.html)
        toonAntwoorden: (idx[k][11] !== false && String(idx[k][11]).toUpperCase() !== "FALSE"), // Check voor toon antwoorden instelling
        videoUrl: idx[k][12] ? String(idx[k][12]).trim() : "",
        videoDuur: parseInt(idx[k][13], 10) || 0
      }; 
      break; 
    }
  }
  
  if(!meta) {
    throw new Error("Toets niet gevonden in database.");
  }

  var shDb = ss.getSheetByName(TAB_DB);
  
  if(!shDb) {
    throw new Error("Database vragen niet gevonden.");
  }
  
  var db = shDb.getDataRange().getValues();
  var dbDisplay = shDb.getDataRange().getDisplayValues(); // TOEGEVOEGD: Haal display waardes op om percentages te behouden
  var qs = [];
  
  function getImageSrc_(urlOrLink) {
    if (!urlOrLink) {
      return "";
    }
    var s = String(urlOrLink).trim();
    if (s.indexOf("data:") === 0) {
      return s;
    }
    var m = s.match(/[-\w]{25,}/);
    if (m) {
      try {
        var file = DriveApp.getFileById(m[0]);
        var blob = file.getBlob();
        var mime = blob.getContentType() || 'image/png';
        return "data:" + mime + ";base64," + Utilities.base64Encode(blob.getBytes());
      } catch(e) {
        return "https://lh3.googleusercontent.com/d/" + m[0];
      }
    }
    return s;
  }

  for(var r = 1; r < db.length; r++) {
    if(db[r][0] != toetsId) {
      continue;
    }
    qs.push({
      id: r + 1,
      vraag: String(dbDisplay[r][1] || ""), 
      afbeelding: getImageSrc_(db[r][2]),
      opties: [dbDisplay[r][3], dbDisplay[r][4], dbDisplay[r][5], dbDisplay[r][6]].filter(function(v) { return v !== ""; }),
      multipleAnswers: String(normalizeCorrectIdx_(db[r][7])).indexOf(",") !== -1,
      verplicht: db[r][8] === true, 
      uitleg: String(dbDisplay[r][9] || "")
    });
  }
  
  var v = qs.filter(function(q) { return q.verplicht; });
  var o = qs.filter(function(q) { return !q.verplicht; }).sort(function() { return 0.5 - Math.random(); });
  
  // Limiet notatie: Als aantal 0 is of leeg, pak dan alle vragen
  var finalQuestions = [];
  if (!meta.aantal || meta.aantal === 0 || meta.aantal >= (v.length + o.length)) {
    finalQuestions = v.concat(o);
  } else {
    var take = Math.max(0, meta.aantal - v.length);
    finalQuestions = v.concat(o.slice(0, take));
  }
  
  finalQuestions.sort(function() { return 0.5 - Math.random(); });

  var attemptsSheet = getQuizAttemptsSheet_(ss);
  var attemptsRows = attemptsSheet.getDataRange().getValues();
  var attemptRowNumber = -1;
  var attempt = null;
  for (var a = attemptsRows.length - 1; a >= 1; a--) {
    var sameAttempt = String(attemptsRows[a][1]) === String(toetsId) &&
      String(attemptsRows[a][2] || "") === String(sid || "") &&
      (sid || String(attemptsRows[a][3] || "").toLowerCase() === email) &&
      attemptsRows[a][9] === "OPEN";
    if (sameAttempt) {
      try {
        attempt = {
          token: String(attemptsRows[a][0]),
          toetsId: String(attemptsRows[a][1]),
          sid: String(attemptsRows[a][2] || ""),
          email: String(attemptsRows[a][3] || ""),
          questionIds: JSON.parse(attemptsRows[a][4] || "[]"),
          norm: attemptsRows[a][5],
          startedAt: attemptsRows[a][6],
          maxTijd: attemptsRows[a][7],
          invited: attemptsRows[a][8] === true
        };
        attemptRowNumber = a + 1;
      } catch(e) {
        attempt = null;
      }
      break;
    }
  }

  var attemptToken = attempt ? attempt.token : "";
  if (attempt && attempt.toetsId === String(toetsId) && attempt.sid === (sid || "")) {
    var questionsById = {};
    qs.forEach(function(q) { questionsById[String(q.id)] = q; });
    var cachedQuestions = attempt.questionIds.map(function(id) { return questionsById[String(id)]; });
    if (cachedQuestions.every(Boolean)) {
      finalQuestions = cachedQuestions;
      startedAt = new Date(attempt.startedAt);
      maxTijd = attempt.maxTijd;
      meta.norm = attempt.norm;
    } else {
      attempt = null;
    }
  } else {
    attempt = null;
  }

  if (!attempt) {
    attemptToken = Utilities.getUuid().replace(/-/g, "");
    attempt = {
      token: attemptToken,
      toetsId: String(toetsId),
      sid: sid || "",
      email: email,
      questionIds: finalQuestions.map(function(q) { return q.id; }),
      norm: meta.norm,
      startedAt: startedAt,
      maxTijd: maxTijd,
      invited: !!sid
    };
    var attemptRow = [
      attempt.token,
      attempt.toetsId,
      attempt.sid,
      attempt.email,
      JSON.stringify(attempt.questionIds),
      attempt.norm,
      attempt.startedAt,
      attempt.maxTijd,
      attempt.invited,
      "OPEN"
    ];
    if (attemptRowNumber > 0) {
      attemptsSheet.getRange(attemptRowNumber, 1, 1, attemptRow.length).setValues([attemptRow]);
    } else {
      attemptsSheet.appendRow(attemptRow);
    }
  }

  return { meta: meta, vragen: finalQuestions, attemptToken: attemptToken };
}

function submitQuizResult(res, tid) {
  try {
    if (!res || typeof res !== "object") {
      throw new Error("De toetsinzending is ongeldig.");
    }
    var clientStartTijd = res.startTijd;

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var attemptToken = String(res.attemptToken || "");
    var attemptsSheet = getQuizAttemptsSheet_(ss);
    var attemptsRows = attemptsSheet.getDataRange().getValues();
    var attemptRowNumber = -1;
    var attempt = null;
    for (var a = attemptsRows.length - 1; a >= 1; a--) {
      if (String(attemptsRows[a][0]) === attemptToken && attemptsRows[a][9] === "OPEN") {
        attempt = {
          token: String(attemptsRows[a][0]),
          toetsId: String(attemptsRows[a][1]),
          sid: String(attemptsRows[a][2] || ""),
          email: String(attemptsRows[a][3] || ""),
          questionIds: JSON.parse(attemptsRows[a][4] || "[]"),
          norm: attemptsRows[a][5],
          startedAt: attemptsRows[a][6],
          maxTijd: attemptsRows[a][7],
          invited: attemptsRows[a][8] === true
        };
        attemptRowNumber = a + 1;
        break;
      }
    }
    if (!attempt) {
      throw new Error("Deze toetssessie is niet meer beschikbaar. Laad de toets opnieuw en probeer het nogmaals.");
    }

    if (attempt.toetsId !== String(tid) || attempt.sid !== String(res.sid || "") || attempt.invited !== !!res.sid) {
      throw new Error("De toetsinzending hoort niet bij deze toets.");
    }

    var shDb = ss.getSheetByName(TAB_DB);
    if (!shDb) throw new Error("Database vragen niet gevonden.");
    var dbData = shDb.getDataRange().getValues();
    var dbDisplay = shDb.getDataRange().getDisplayValues();
    var submittedAnswers = Array.isArray(res.answers) ? res.answers : [];
    var answerMap = {};
    submittedAnswers.forEach(function(answer) {
      var id = String(answer.questionId || "");
      if (!id || answerMap[id] || !Array.isArray(answer.values)) {
        throw new Error("De antwoorden van de toetsinzending zijn ongeldig.");
      }
      answerMap[id] = answer.values.map(function(value) { return String(value); });
    });

    if (!attempt.questionIds.length || submittedAnswers.length !== attempt.questionIds.length) {
      throw new Error("Niet alle toetsvragen zijn beantwoord.");
    }

    var details = [];
    var correctCount = 0;
    attempt.questionIds.forEach(function(questionId) {
      var id = String(questionId);
      var rowNumber = parseInt(id, 10);
      var row = dbData[rowNumber - 1];
      var displayRow = dbDisplay[rowNumber - 1];
      var selected = answerMap[id];
      if (!row || String(row[0]) !== String(tid) || !selected || selected.length === 0) {
        throw new Error("De toetsvragen zijn gewijzigd of de inzending is ongeldig.");
      }

      var options = [displayRow[3], displayRow[4], displayRow[5], displayRow[6]].filter(function(value) { return value !== ""; });
      var normalizedSelected = selected.map(function(value) {
        if (!/^[1-4]$/.test(value) || Number(value) > options.length) {
          throw new Error("Een antwoordkeuze is ongeldig.");
        }
        return value;
      }).sort();
      if (normalizedSelected.some(function(value, index) { return index > 0 && value === normalizedSelected[index - 1]; })) {
        throw new Error("Een antwoordkeuze is dubbel ingediend.");
      }

      var correctIndices = String(normalizeCorrectIdx_(row[7])).split(",").map(function(value) {
        return value.trim();
      }).filter(Boolean).sort();
      var isCorrect = normalizedSelected.join(",") === correctIndices.join(",");
      if (isCorrect) correctCount++;

      details.push({
        question: String(displayRow[1] || ""),
        givenAnswer: normalizedSelected.map(function(value) { return options[Number(value) - 1]; }).join(", "),
        correctAnswer: correctIndices.map(function(value) { return options[Number(value) - 1]; }).join(", "),
        isCorrect: isCorrect,
        uitleg: String(displayRow[9] || "")
      });
    });

    if (Object.keys(answerMap).length !== attempt.questionIds.length) {
      throw new Error("De toetsinzending bevat onbekende vragen.");
    }

    var norm = parseNorm(attempt.norm);
    var score = Math.round((correctCount / attempt.questionIds.length) * 100);
    var startTijd = new Date(attempt.startedAt);
    if (isNaN(startTijd.getTime())) throw new Error("De starttijd van de toets is ongeldig.");
    var stat = score >= norm ? "Geslaagd" : "Gezakt";
    var userEmail = attempt.email || "Onbekende Gebruiker";
    res.score = score;
    res.correct = correctCount;
    res.total = attempt.questionIds.length;
    res.details = details;
    res.norm = norm;
    res.startTijd = startTijd.toISOString();
    res.toetsNaam = "";
    var rawStartTijd = startTijd;
    var ccEmails = "";
    var managerEmail = "";
    var herhaling = "";
    var maxTijd = Number(attempt.maxTijd) || 0;
    var aantalVragen = res.total;
    
    var shIdx = ss.getSheetByName(TAB_INDEX);
    var stuurManagerMail = true;
    var toonAntwoorden = true;
    var toetsBestaat = false;
    if (shIdx) {
        var idxRows = shIdx.getDataRange().getValues();
        var searchTid = String(tid).trim();
        for (var m = 1; m < idxRows.length; m++) {
            if (String(idxRows[m][0]).trim() === searchTid) {
                toetsBestaat = true;
                res.toetsNaam = idxRows[m][1] || "-";
                stuurManagerMail = (idxRows[m][10] !== false && String(idxRows[m][10]).toUpperCase() !== "FALSE");
                toonAntwoorden = (idxRows[m][11] !== false && String(idxRows[m][11]).toUpperCase() !== "FALSE");
                break;
            }
        }
    }
    if (!toetsBestaat) throw new Error("Toets niet gevonden.");
    
    if(res.sid) {
      var shS = ss.getSheetByName(TAB_SESSIES);
      var sessionFound = false;
      if (shS) {
        var sD = shS.getDataRange().getValues();
        for(var i = sD.length - 1; i >= 1; i--) {
          if(String(sD[i][9]) === res.sid) {
            if (sD[i][10] === "USED" || sD[i][10] !== "BEZIG") {
              throw new Error("Deze toetslink is al gebruikt of niet gestart.");
            }
            var inviteEmail = sD[i][8]; 
            ccEmails = sD[i][3];
            managerEmail = sD[i][4];
            herhaling = sD[i][7];
            maxTijd = parseFloat(String(sD[i][6]).replace(',', '.')) || 0;
            aantalVragen = sD[i][12];
            rawStartTijd = sD[i][11];
            if (rawStartTijd instanceof Date) startTijd = rawStartTijd;
            
            if (inviteEmail && String(inviteEmail).indexOf("@") > -1) {
              userEmail = inviteEmail; 
            }
            
            shS.getRange(i+1, 11).setValue("USED");
            sessionFound = true;
            break;
          }
        }
      }
      if (!sessionFound) throw new Error("De toetslink is ongeldig.");
    } else {
      var shExistingResults = ss.getSheetByName(TAB_RESULTATEN);
      if (userEmail.indexOf("@") > -1 && shExistingResults) {
        var existingRows = shExistingResults.getDataRange().getValues();
        var alreadySubmitted = existingRows.slice(1).some(function(row) {
          return String(row[0]).toLowerCase().indexOf(userEmail.toLowerCase()) !== -1 && String(row[7]) === String(tid);
        });
        if (alreadySubmitted) throw new Error("Je hebt deze toets al gemaakt.");
      }
    }

    var completedAt = new Date();
    var recordedStartTijd = new Date(clientStartTijd);
    var clientDuurSec = parseInt(res.duurSec, 10);
    
    // Klok-gesynchroniseerde starttijd en duur berekening (ongevoelig voor klokafwijking tussen client en server)
    if (!isNaN(clientDuurSec) && clientDuurSec >= 0) {
      recordedStartTijd = new Date(completedAt.getTime() - (clientDuurSec * 1000));
    } else if (isNaN(recordedStartTijd.getTime()) || recordedStartTijd.getTime() < (startTijd.getTime() - 3600000) || recordedStartTijd.getTime() > (completedAt.getTime() + 300000)) {
      recordedStartTijd = startTijd;
    }

    res.startTijd = recordedStartTijd.toISOString();
    res.duur = Math.round((completedAt.getTime() - recordedStartTijd.getTime()) / 600) / 100;
    if (res.duur < 0) res.duur = 0;
    
    if (maxTijd > 0) {
      var secBezig = Math.round((completedAt.getTime() - recordedStartTijd.getTime()) / 1000);
      // Ruime buffer van 60 seconden voor netwerklatentie en trage verbindingen
      if (secBezig > (maxTijd * 60) + 60) {
        stat = "Gezakt (Tijd overschreden)";
      }
    }
    
    var managerNaamInput = res.kandidaatNaam || "";
    var displayKandidaat = userEmail;
    if (managerNaamInput && String(managerNaamInput).trim() !== "") {
       displayKandidaat = managerNaamInput + " (" + userEmail + ")";
    }

    var toetsNaam = res.toetsNaam;
    var duur = res.duur;
    var shRes = ss.getSheetByName(TAB_RESULTATEN);
    
    if(!shRes) {
      shRes = ss.insertSheet(TAB_RESULTATEN);
    }
    
    var resRow = [
      displayKandidaat, 
      toetsNaam, 
      new Date(res.startTijd), 
      new Date(), 
      duur,
      res.correct + "/" + res.total + " " + res.score + "%", 
      stat, 
      tid, 
      JSON.stringify(res.details), // Data in database heeft altijd antwoorden!
      res.sid || ""
    ];
    
    shRes.insertRowBefore(2);
    shRes.getRange(2, 1, 1, resRow.length).setValues([resRow]);
    attemptsSheet.deleteRow(attemptRowNumber);
    
    if (userEmail && userEmail.indexOf("@") > -1) {
      res.userEmail = displayKandidaat; 
      res.finalStatus = stat;
      res.norm = norm;
      res.toetsNaam = toetsNaam;
      res.datum = new Date().toLocaleDateString('nl-NL');
      res.toonAntwoorden = toonAntwoorden; // Doorgeven aan PDF generator
      
      var sendToMail = userEmail;
      try {
        if (typeof getEmployeeMapping_ === 'function') {
          var empMap = getEmployeeMapping_();
          var emailClean = userEmail.toLowerCase().trim();
          if (empMap[emailClean]) {
            var isGL = empMap[emailClean].isGroepsleider;
            var privMail = empMap[emailClean].privateEmail;
            if (!isGL && privMail && privMail.indexOf("@") > -1 && privMail !== emailClean) {
              sendToMail = privMail;
            }
          }
        }
      } catch(e) {}
      
      res.sendToEmail = sendToMail; 
      
      try { 
        sendPDFEmail(res); 
      } catch(mailErr) { 
        Logger.log("Kandidaat mail fout: " + mailErr); 
      }
    }
    
    var recipients = [];
    
    // OPLOSSING: De manager wordt alléén gemaild als de stuurManagerMail toggle op True staat
    if(managerEmail && stuurManagerMail) {
      recipients.push(managerEmail);
    }
    
    if(ccEmails) { 
      var splits = ccEmails.split(/[\n,;]+/); 
      for(var k = 0; k < splits.length; k++) {
        recipients.push(splits[k].trim()); 
      }
    }
    
    var uCC = []; 
    for(var l = 0; l < recipients.length; l++) { 
      if(recipients[l].indexOf("@") > -1 && uCC.indexOf(recipients[l]) == -1) {
        uCC.push(recipients[l]); 
      }
    }

    if(uCC.length > 0) {
      var appUrl = ScriptApp.getService().getUrl();
      appUrl = appUrl.replace(/\/dev$/i, "/exec");
      var logoUrl = CONFIG.LOGO_URL;
      
      var body = `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f6f9; padding: 20px; border-radius: 8px;">
          <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; border-top: 4px solid #D32F2F; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px;">
              <tr>
                <td style="vertical-align: middle;"><h2 style="color: #232323; font-size: 20px; margin: 0;">Notificatie Uitslag Toets</h2></td>
                <td style="text-align: right; vertical-align: middle;"><img src="${logoUrl}" alt="G4S Logo" style="height: 50px;"></td>
              </tr>
            </table>
            <p style="color: #333; font-size: 15px; line-height: 1.6;">Beste beheerder/manager,</p>
            <p style="color: #333; font-size: 15px; line-height: 1.6;">Er is zojuist een toets afgerond door <strong>${displayKandidaat}</strong>.</p>
            <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0; border: 1px solid #eee;">
              <p style="margin: 0 0 8px 0; color: #333;"><strong>Toets:</strong> ${toetsNaam}</p>
              <p style="margin: 0 0 8px 0; color: #333;"><strong>Score:</strong> ${res.score}%</p>
              <p style="margin: 0; color: #333;"><strong>Resultaat:</strong> ${stat}</p>
            </div>
            <div style="text-align: center; margin: 30px 0;">
              <a href="${appUrl}" style="background-color: #D32F2F; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; text-transform: uppercase;">BEKIJK IN PORTAL</a>
            </div>
            <p style="color: #666; font-size: 13px; line-height: 1.5; margin-top: 30px;">
              Met vriendelijke groet,<br><strong>G4S Operation Support Team</strong>
            </p>
          </div>
        </div>
      `;
      try { 
        MailApp.sendEmail({ to: uCC.join(","), subject: "Notificatie: " + displayKandidaat + " - " + stat, htmlBody: body }); 
      } catch(ccErr) {}
    }

    var dagenHerhaling = parseInt(herhaling, 10);
    if(stat == "Geslaagd" && !isNaN(dagenHerhaling) && dagenHerhaling > 0 && (!managerNaamInput || String(managerNaamInput).trim() === "")) {
      var d = new Date(); 
      d.setDate(d.getDate() + dagenHerhaling);
      
      var hSheet = ss.getSheetByName(TAB_HERINNERINGEN);
      if (!hSheet) { 
        hSheet = ss.insertSheet(TAB_HERINNERINGEN); 
        hSheet.appendRow(["Email", "ToetsID", "Datum", "Status", "ManagerEmail", "CC", "Norm", "AantalVragen", "Tijd", "HerhalingDagen"]); 
      }
      hSheet.appendRow([userEmail, tid, d, "Gepland", managerEmail, ccEmails, norm, (aantalVragen || res.total), (maxTijd || 0), dagenHerhaling]);
    }
    
    try { clearSystemCaches(); } catch(cacheErr) {}
    
    return {
      finalStatus: stat,
      norm: norm,
      score: score,
      correct: correctCount,
      total: attempt.questionIds.length,
      details: toonAntwoorden ? details : details.map(function(detail) {
        return {
          question: detail.question,
          givenAnswer: detail.givenAnswer,
          isCorrect: detail.isCorrect,
          uitleg: detail.uitleg
        };
      })
    };
  } catch(e) { 
    throw new Error("Fout bij opslaan resultaat: " + e.message); 
  }
}

function updateResultaatStatus(toetsId, email, sid, nieuweStatus) {
  try {
    var allowedIds = getAllowedToetsIds_();
    if (!allowedIds._isAdmin && !allowedIds[toetsId]) {
      throw new Error("Geen rechten om dit resultaat te bewerken.");
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shRes = ss.getSheetByName(TAB_RESULTATEN);
    var data = shRes.getDataRange().getValues();
    
    for (var j = data.length - 1; j >= 1; j--) {
      var rowEmail = String(data[j][0]).toLowerCase();
      var isMatch = false;
      
      if (rowEmail.indexOf(email.toLowerCase()) > -1 && data[j][7] == toetsId && data[j][9] === sid) {
        isMatch = true;
      }
      
      if (isMatch) {
        shRes.getRange(j + 1, 7).setValue(nieuweStatus); 
        try { clearSystemCaches(); } catch(e){}
        return { success: true, nieuweStatus: nieuweStatus };
      }
    }
    throw new Error("Resultaat niet gevonden in database.");
  } catch(e) {
    throw new Error("Fout bij updaten status: " + e.message);
  }
}