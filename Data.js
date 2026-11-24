// ====================================================================================================
// G4S KENNISTOETSEN - DATA BEHEER, EDITOR, UITNODIGINGEN & LOGGING
// ====================================================================================================

// Let op: Globale variabelen (zoals TAB_ namen en CONFIG) worden nu centraal vanuit Code.gs ingeladen.

// ----------------------------------------------------------------------------------------------------
// 0. AUTORISATIE HULPMIDDEL
// ----------------------------------------------------------------------------------------------------

function autoriseerApp() {
  DriveApp.getRootFolder();
  Logger.log("App is succesvol geautoriseerd voor Google Drive!");
}

// ----------------------------------------------------------------------------------------------------
// 1. LOGGING & AUDIT
// ----------------------------------------------------------------------------------------------------

function writeLog_(action, details) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(TAB_MUTATIE);
    
    if (!sh) {
      sh = ss.insertSheet(TAB_MUTATIE);
      sh.appendRow(["Datum", "Gebruiker", "Actie", "Details"]);
    }
    
    var email = "Systeem";
    try { 
      email = Session.getActiveUser().getEmail() || "Systeem"; 
    } catch(e) {}
    
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, 4).setValues([[new Date(), email, action, details]]);
  } catch (e) { 
    Logger.log("Loggen mislukt: " + e.message); 
  }
}

function getAuditLogs() {
  var allowed = getAllowedToetsIds_();
  
  if (!allowed._isAdmin) {
    throw new Error("Alleen beheerders mogen de logs inzien.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_MUTATIE);
  
  if (!sh) {
    return [];
  }
  
  var data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  
  return data.slice(1).map(function(r) {
    var d = r[0];
    return { 
      datum: (d instanceof Date) ? d.toLocaleString('nl-NL') : d, 
      user: r[1], 
      actie: r[2], 
      details: r[3] 
    };
  });
}

// ----------------------------------------------------------------------------------------------------
// 2. RESULTAAT & TOETS BEHEER (VERWIJDEREN)
// ----------------------------------------------------------------------------------------------------

function deleteResult(resData) {
  var allowed = getAllowedToetsIds_();
  var email = Session.getActiveUser().getEmail().toLowerCase();
  
  var isManager = false;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var shIdx = ss.getSheetByName(TAB_INDEX);
  var idxRows = shIdx.getDataRange().getValues();
  
  for (var i = 1; i < idxRows.length; i++) {
    if (idxRows[i][0] == resData.toetsId) {
      var mgrEmail = String(idxRows[i][6]).toLowerCase();
      var splits = mgrEmail.split(/[\n,;]+/);
      for (var s = 0; s < splits.length; s++) {
        if (splits[s].trim() === email) {
          isManager = true;
          break;
        }
      }
      break;
    }
  }

  if (!allowed._isAdmin && !isManager) {
    throw new Error("Je hebt geen rechten om dit resultaat te verwijderen.");
  }

  var shRes = ss.getSheetByName(TAB_RESULTATEN);
  var data = shRes.getDataRange().getValues();
  
  for (var j = data.length - 1; j >= 1; j--) {
    var rowEmail = String(data[j][0]).toLowerCase();
    var rowDate = (data[j][3] instanceof Date) ? data[j][3].toISOString() : String(data[j][3]);
    var targetDate = resData.datumIso;

    if (rowEmail === resData.email.toLowerCase() && (rowDate === targetDate || data[j][9] === resData.sid)) {
      shRes.deleteRow(j + 1);
      writeLog_("RESULTAAT_VERWIJDERD", "Deelnemer: " + resData.email + " - Toets: " + resData.toetsNaam);
      clearSystemCaches(); 
      return { success: true };
    }
  }
  
  throw new Error("Resultaat niet gevonden in de database.");
}

function updateResultaatStatus(toetsId, email, sid, nieuweStatus) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var shRes = ss.getSheetByName(TAB_RESULTATEN);
  if (!shRes) return;
  
  var data = shRes.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] == toetsId && data[i][0] == email && data[i][9] == sid) {
      shRes.getRange(i+1, 7).setValue(nieuweStatus);
      writeLog_("SCORE_GEFORCEERD", "Resultaat van " + email + " voor toets " + toetsId + " aangepast naar: " + nieuweStatus);
      clearSystemCaches();
      return { success: true };
    }
  }
  throw new Error("Resultaat niet gevonden.");
}

function deleteToets(toetsId) {
  var allowed = getAllowedToetsIds_();
  
  if (!allowed._isAdmin) {
    throw new Error("Alleen beheerders mogen toetsen definitief verwijderen.");
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  var shIdx = ss.getSheetByName(TAB_INDEX);
  var idxData = shIdx.getDataRange().getValues();
  var toetsNaam = "Onbekende Toets";
  
  for (var i = idxData.length - 1; i >= 1; i--) {
    if (idxData[i][0] == toetsId) {
      toetsNaam = idxData[i][1];
      shIdx.deleteRow(i + 1);
    }
  }
  
  var shDb = ss.getSheetByName(TAB_DB);
  var dbData = shDb.getDataRange().getValues();
  
  for (var j = dbData.length - 1; j >= 1; j--) {
    if (dbData[j][0] == toetsId) {
      var imgUrl = dbData[j][2];
      
      if (imgUrl && String(imgUrl).indexOf("drive.google.com") > -1) {
        deleteImageFromDrive(imgUrl);
      }
      
      shDb.deleteRow(j + 1);
    }
  }
  
  writeLog_("TOETS_VERWIJDERD", "Toets: " + toetsNaam + " (ID: " + toetsId + ")");
  clearSystemCaches(); 
  return { success: true };
}

// ----------------------------------------------------------------------------------------------------
// 3. DROPDOWNS & TOEGANG & ROUTING
// ----------------------------------------------------------------------------------------------------

function getEditorDropdownData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var regios = [];
    var klanten = [];

    var shReg = ss.getSheetByName(TAB_REGIO);
    if (shReg) {
      var dataReg = shReg.getDataRange().getValues();
      var rSet = new Set();
      for (var i = 1; i < dataReg.length; i++) {
        if (dataReg[i][3]) {
          rSet.add(String(dataReg[i][3]).trim());
        }
      }
      regios = Array.from(rSet).sort();
    }

    var shKlant = ss.getSheetByName(TAB_KLANTEN);
    if (shKlant) {
      var dataKlant = shKlant.getDataRange().getValues();
      var kSet = new Set();
      for (var j = 1; j < dataKlant.length; j++) {
        if (dataKlant[j][1]) {
          kSet.add(String(dataKlant[j][1]).trim());
        }
      }
      klanten = Array.from(kSet).sort();
    }
    
    return { regios: regios, klanten: klanten };
  } catch (e) { 
    return { regios: [], klanten: [] }; 
  }
}

function getCurrentUserEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
  } catch(e) {
    return "";
  }
}

function getViewAsEmail_(requestedEmail) {
  var currentEmail = getCurrentUserEmail_();
  var requested = String(requestedEmail || "").toLowerCase().trim();
  if (requested && requested !== currentEmail && getAllowedToetsIds_()._isAdmin) {
    return requested;
  }
  return currentEmail;
}

function getAllowedToetsIds_(viewAsEmail) {
  var email = getCurrentUserEmail_();
  var requestedEmail = String(viewAsEmail || "").toLowerCase().trim();
  if (requestedEmail && requestedEmail !== email && getAllowedToetsIds_()._isAdmin) {
    email = requestedEmail;
  }
  
  var allowed = { _isAdmin: false, canCreate: false };
  if (!email) {
    return allowed;
  }

  if (typeof CONFIG !== 'undefined' && (email === CONFIG.SUPPORT_EMAIL || email === CONFIG.SUPPORT_EMAIL_ALT)) {
    allowed._isAdmin = true;
    allowed.canCreate = true;
    return allowed;
  } else if (email === "operation.supportteam@nl.g4s.com" || email === "operation.supporteam@nl.g4s.com") {
    allowed._isAdmin = true;
    allowed.canCreate = true;
    return allowed;
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var shRechten = ss.getSheetByName(TAB_RECHTEN);
  if (shRechten) {
    var rData = shRechten.getDataRange().getValues();
    for (var r = 1; r < rData.length; r++) {
      if (String(rData[r][0]).toLowerCase().trim() === email) {
        var rol = String(rData[r][1]).toLowerCase().trim();
        if (rol === "systeembeheerder") {
          allowed._isAdmin = true;
          allowed.canCreate = true;
          return allowed;
        } else if (rol === "toetsenmaker") {
          allowed.canCreate = true;
        }
      }
    }
  }

  var shIdx = ss.getSheetByName(TAB_INDEX);
  if (shIdx) {
    var idxData = shIdx.getDataRange().getValues();
    for (var i = 1; i < idxData.length; i++) {
      var mgrEmail = String(idxData[i][6]).toLowerCase();
      var splits = mgrEmail.split(/[\n,;]+/);
      for (var s = 0; s < splits.length; s++) {
        if (splits[s].trim() === email) {
          allowed[idxData[i][0]] = true;
          break; 
        }
      }
    }
  }
  
  return allowed;
}

// ----------------------------------------------------------------------------------------------------
// GEBRUIKERSBEHEER FUNCTIES
// ----------------------------------------------------------------------------------------------------

function getGebruikers() {
  var allowed = getAllowedToetsIds_();
  if (!allowed._isAdmin) throw new Error("Geen rechten om gebruikers te bekijken.");

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_RECHTEN);
  if (!sh) return [];

  var data = sh.getDataRange().getValues();
  var res = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      res.push({ email: data[i][0], rol: data[i][1] });
    }
  }
  return res;
}

function saveGebruiker(email, rol) {
  var allowed = getAllowedToetsIds_();
  if (!allowed._isAdmin) throw new Error("Geen rechten om gebruikers toe te voegen.");

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_RECHTEN);
  if (!sh) {
    sh = ss.insertSheet(TAB_RECHTEN);
    sh.appendRow(["Email", "Rol"]);
  }

  email = String(email).toLowerCase().trim();
  var data = sh.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === email) {
      sh.getRange(i+1, 2).setValue(rol);
      writeLog_("GEBRUIKER_GEWIJZIGD", "Email: " + email + ", Nieuwe rol: " + rol);
      clearSystemCaches();
      return {success: true};
    }
  }
  
  sh.appendRow([email, rol]);
  writeLog_("GEBRUIKER_AANGEMAAKT", "Email: " + email + ", Rol: " + rol);
  clearSystemCaches();
  return {success: true};
}

function deleteGebruiker(email) {
  var allowed = getAllowedToetsIds_();
  if (!allowed._isAdmin) throw new Error("Geen rechten om gebruikers te verwijderen.");

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_RECHTEN);
  if (!sh) return {success: false};

  email = String(email).toLowerCase().trim();
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).toLowerCase().trim() === email) {
      sh.deleteRow(i+1);
      writeLog_("GEBRUIKER_VERWIJDERD", "Email: " + email);
    }
  }
  clearSystemCaches();
  return {success: true};
}

// ----------------------------------------------------------------------------------------------------
// GEDEELDE ACCOUNTS BEHEER (ACCOUNT AANVULLING)
// ----------------------------------------------------------------------------------------------------

function getSharedAccounts() {
  var allowed = getAllowedToetsIds_();
  if (!allowed._isAdmin) throw new Error("Geen rechten om gedeelde accounts te bekijken.");

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_AANVULLING);
  if (!sh) return [];

  var data = sh.getDataRange().getValues();
  var res = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][4]) { 
      res.push({
        unit: data[i][0] || "",
        kp: data[i][1] || "",
        bu: data[i][2] || "",
        regio: data[i][3] || "",
        email: data[i][4] || ""
      });
    }
  }
  return res;
}

function saveSharedAccount(acc) {
  var allowed = getAllowedToetsIds_();
  if (!allowed._isAdmin) throw new Error("Geen rechten om gedeelde accounts toe te voegen.");

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_AANVULLING);
  if (!sh) {
    sh = ss.insertSheet(TAB_AANVULLING);
    sh.appendRow(["Unit", "KP", "Businessunit", "Regio", "E-mail adres"]);
  }

  var email = String(acc.email).toLowerCase().trim();
  var data = sh.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).toLowerCase().trim() === email) {
      sh.getRange(i+1, 1, 1, 5).setValues([[acc.unit, acc.kp, acc.bu, acc.regio, email]]);
      writeLog_("ACCOUNT_AANVULLING_GEWIJZIGD", "Email: " + email);
      clearSystemCaches();
      return {success: true};
    }
  }
  
  sh.appendRow([acc.unit, acc.kp, acc.bu, acc.regio, email]);
  writeLog_("ACCOUNT_AANVULLING_AANGEMAAKT", "Email: " + email);
  clearSystemCaches();
  return {success: true};
}

function deleteSharedAccount(email) {
  var allowed = getAllowedToetsIds_();
  if (!allowed._isAdmin) throw new Error("Geen rechten om gedeelde accounts te verwijderen.");

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TAB_AANVULLING);
  if (!sh) return {success: false};

  email = String(email).toLowerCase().trim();
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][4]).toLowerCase().trim() === email) {
      sh.deleteRow(i+1);
      writeLog_("ACCOUNT_AANVULLING_VERWIJDERD", "Email: " + email);
    }
  }
  clearSystemCaches();
  return {success: true};
}

function getEmployeeMapping_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var employeeMap = {};
  
  try {
    var shEmp = ss.getSheetByName(TAB_EMP);
    if (shEmp) {
      var empData = shEmp.getDataRange().getValues();
      if (empData.length > 0) {
        var upnIdx = 44;  
        var privIdx = 17; 
        var funcIdx = 25; 
        
        var headerRow = empData[0];
        for (var c = 0; c < headerRow.length; c++) {
          var colName = String(headerRow[c]).toLowerCase().replace(/[^a-z0-9]/g, '');
          if (colName === "g4supn" || colName === "upn") upnIdx = c;
          if (colName === "email") privIdx = c;
          if (colName === "functie") funcIdx = c;
        }
        
        for (var r = 1; r < empData.length; r++) {
          var rawUpn = empData[r].length > upnIdx ? empData[r][upnIdx] : "";
          var upn = String(rawUpn || "").toLowerCase().trim();
          
          var rawPriv = empData[r].length > privIdx ? empData[r][privIdx] : "";
          var privMail = String(rawPriv || "").toLowerCase().trim();
          
          var rawFunc = empData[r].length > funcIdx ? empData[r][funcIdx] : "";
          var functie = String(rawFunc || "").toLowerCase().trim();
          var isGL = (functie.indexOf("groepsleider") > -1);
          
          if (upn && upn.indexOf("@") > -1) {
            var empObj = { upn: upn, privateEmail: privMail, isGroepsleider: isGL, source: "HR" };
            employeeMap[upn] = empObj;
            
            if (privMail && privMail.indexOf("@") > -1) {
              employeeMap[privMail] = empObj;
            }
          }
        }
      }
    }
    
    var shAanvulling = ss.getSheetByName(TAB_AANVULLING);
    if (shAanvulling) {
      var aanvData = shAanvulling.getDataRange().getValues();
      if (aanvData.length > 0) {
        var colIdx = 4; 
        var h2 = aanvData[0];
        for (var c2 = 0; c2 < h2.length; c2++) {
          var cn2 = String(h2[c2]).toLowerCase().replace(/[^a-z0-9]/g, '');
          if (cn2.indexOf("email") > -1) { colIdx = c2; break; }
        }
        
        for (var r2 = 1; r2 < aanvData.length; r2++) {
          var rawEml = aanvData[r2].length > colIdx ? aanvData[r2][colIdx] : "";
          var eml = String(rawEml || "").toLowerCase().trim();
          
          if (eml && eml.indexOf("@") > -1 && (!employeeMap[eml] || employeeMap[eml].source !== "HR")) {
            employeeMap[eml] = { upn: eml, privateEmail: eml, isGroepsleider: false, source: "SHARED" };
          }
        }
      }
    }
  } catch(e) {
    Logger.log("Fout bij uitlezen werknemers mapping: " + e.message);
  }
  
  return employeeMap;
}

// ----------------------------------------------------------------------------------------------------
// 4. DATA FETCHING & IMPORT & UPLOAD
// ----------------------------------------------------------------------------------------------------

function uploadImageToDrive(base64, filename, mimeType, toetsNaam, isConcept) {
  try {
    var mainFolderId = (typeof CONFIG !== 'undefined' && CONFIG.DRIVE_FOLDER_ID) ? CONFIG.DRIVE_FOLDER_ID : "1CGktm8Hro1kbZVARuMAAkhoJUUk_xQKa";
    var mainFolder = DriveApp.getFolderById(mainFolderId);
    
    var folderName = toetsNaam ? toetsNaam.replace(/[\\/:*?"<>|]/g, "_") : (isConcept ? "Concept_Aanvragen" : "Overige_Toetsen");
    var subFolders = mainFolder.getFoldersByName(folderName);
    var subFolder;
    
    if (subFolders.hasNext()) {
      subFolder = subFolders.next();
    } else {
      subFolder = mainFolder.createFolder(folderName);
    }
    
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, filename);
    var file = subFolder.createFile(blob);
    var uploaderEmail = getCurrentUserEmail_();
    if (uploaderEmail) file.setDescription("G4S_UPLOAD_OWNER:" + uploaderEmail);
    
    return file.getUrl();
  } catch(e) { 
    throw new Error("Afbeelding uploaden naar Drive mislukt: " + e.message); 
  }
}

function deleteImageFromDrive(url, toetsId) {
  try {
    if (!url || String(url).indexOf("drive.google.com") === -1) {
      return { success: false };
    }
    
    var m = String(url).match(/[-\w]{25,}/);
    if (m && m[0]) {
      var fileId = m[0];
      var allowed = getAllowedToetsIds_();
      if (!allowed._isAdmin && toetsId && !allowed[toetsId]) {
        throw new Error("Geen rechten om deze toetsafbeelding te verwijderen.");
      }

      var mainFolderId = (typeof CONFIG !== 'undefined' && CONFIG.DRIVE_FOLDER_ID) ? CONFIG.DRIVE_FOLDER_ID : "1CGktm8Hro1kbZVARuMAAkhoJUUk_xQKa";
      var file = DriveApp.getFileById(fileId);
      var parents = file.getParents();
      var inImageFolder = false;
      var managedFolderName = "";
      while (parents.hasNext()) {
        var parent = parents.next();
        if (parent.getId() === mainFolderId) {
          inImageFolder = true;
          break;
        }
        var grandparents = parent.getParents();
        while (grandparents.hasNext()) {
          if (grandparents.next().getId() === mainFolderId) {
            inImageFolder = true;
            managedFolderName = parent.getName();
            break;
          }
        }
        if (inImageFolder) break;
      }
      if (!inImageFolder) {
        throw new Error("De afbeelding staat niet in de beheerde toetsmap.");
      }
      if (!allowed._isAdmin) {
        var currentEmail = getCurrentUserEmail_();
        var ownUpload = currentEmail && file.getDescription() === "G4S_UPLOAD_OWNER:" + currentEmail;
        if (toetsId) {
          var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
          var shDb = ss.getSheetByName(TAB_DB);
          var dbRows = shDb ? shDb.getDataRange().getValues() : [];
          var belongsToTest = dbRows.some(function(row) {
            return String(row[0]) === String(toetsId) && String(row[2] || "").indexOf(fileId) !== -1;
          });
          var shIdx = ss.getSheetByName(TAB_INDEX);
          var testFolderName = "";
          if (shIdx) {
            var indexRows = shIdx.getDataRange().getValues();
            for (var i = 1; i < indexRows.length; i++) {
              if (String(indexRows[i][0]) === String(toetsId)) {
                testFolderName = String(indexRows[i][1] || "").replace(/[\\/:*?"<>|]/g, "_");
                break;
              }
            }
          }
          if (!belongsToTest && (!testFolderName || managedFolderName !== testFolderName) && !ownUpload) {
            throw new Error("De afbeelding hoort niet bij deze toets.");
          }
        } else if (!ownUpload) {
          throw new Error("De afbeelding hoort niet bij deze toets.");
        }
      }
      file.setTrashed(true);
      return { success: true };
    }
  } catch(e) { 
    Logger.log("Verwijderen afbeelding uit Drive mislukt: " + e.message);
    throw e;
  }
  return { success: false };
}

function getAllResultsData(clientEmail) {
  try {
    var allowedIds = getAllowedToetsIds_(clientEmail);
    var email = getViewAsEmail_(clientEmail);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shIdx = ss.getSheetByName(TAB_INDEX);
    var metaMap = {};
    
    if (shIdx) {
      var idxRows = shIdx.getDataRange().getValues();
      for (var j = 1; j < idxRows.length; j++) {
        metaMap[idxRows[j][0]] = { 
          name: idxRows[j][1], 
          norm: parseNorm(idxRows[j][7]), 
          manager: String(idxRows[j][6]).toLowerCase() 
        };
      }
    }

    var shRes = ss.getSheetByName(TAB_RESULTATEN);
    if (!shRes) {
      return [];
    }
    
    var data = shRes.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    
    var rows = data.slice(1).filter(function(r) { 
      return r[0] && String(r[0]).indexOf("@") > -1 && (allowedIds._isAdmin || allowedIds[r[7]]);
    });

    var output = rows.map(function(r, idx) {
      var d = r[3];
      var tId = r[7];
      var tMeta = metaMap[tId] || { name: r[1] || "Onbekende Toets", norm: 70, manager: "" };
      var scObj = parseScoreStr_(r[5]);
      
      var detailsRaw = r[8] || "[]";

      var datumWeergave = "-";
      if (d instanceof Date) {
        datumWeergave = d.toLocaleDateString('nl-NL') + " " + d.toLocaleTimeString('nl-NL', {hour: '2-digit', minute: '2-digit'});
      }

      return {
        id: idx, 
        email: r[0], 
        datum: (d instanceof Date) ? d.toISOString().split('T')[0] : "", 
        datumIso: (d instanceof Date) ? d.toISOString() : d,
        displayDatum: datumWeergave,
        score: scObj.percent, 
        correct: scObj.correct, 
        total: scObj.total, 
        status: r[6],
        toetsId: tId || "old", 
        toetsNaam: tMeta.name, 
        norm: tMeta.norm, 
        details: detailsRaw, 
        sid: r[9],
        canDelete: (allowedIds._isAdmin || tMeta.manager === email)
      };
    }).reverse();
    
    return output;
  } catch(e) { 
    throw new Error("Resultaten laden mislukt: " + e.message); 
  }
}

function getDashboardData(clientEmail, forceRefresh) {
  var email = getViewAsEmail_(clientEmail);
  
  try {
    var url = ScriptApp.getService().getUrl();
    if (url && url.indexOf("/dev") === -1) {
      PropertiesService.getScriptProperties().setProperty('SAVED_WEBAPP_URL', url);
    }
  } catch(e) {}

  if (forceRefresh) {
    clearSystemCaches();
  }
  var version = getGlobalCacheVersion();
  
  var cacheKey = 'dash_' + Utilities.base64Encode(email).substring(0,25) + '_' + version;
  
  var cached = getChunkedCache(cacheKey);
  if (cached) return cached;
  
  var data = getDashboardData_calc_(email);
  putChunkedCache(cacheKey, data, 21600); 
  return data;
}

function getDashboardData_calc_(email) {
  try {
    var allowedIds = getAllowedToetsIds_(email);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shIdx = ss.getSheetByName(TAB_INDEX);
    var toetsNamen = {};
    
    if (shIdx) {
      var idxRows = shIdx.getDataRange().getValues();
      for (var j = 1; j < idxRows.length; j++) {
        toetsNamen[idxRows[j][0]] = idxRows[j][1];
      }
    }

    var baseStats = function() { 
      return { totaal: 0, bezig: 0, verwerkt: 0, geslaagd: 0, gezakt: 0 }; 
    };
    
    var all = baseStats();
    var perToets = {};

    var shRes = ss.getSheetByName(TAB_RESULTATEN);
    if (shRes) {
      var lastRow = shRes.getLastRow();
      if (lastRow > 1) {
        var rows = shRes.getRange(2, 1, lastRow - 1, 8).getValues().filter(function(r) { 
          return r[0] && String(r[0]).trim() !== ""; 
        });
        
        for (var i = 0; i < rows.length; i++) {
          var tId = rows[i][7];
          
          if (!allowedIds._isAdmin && !allowedIds[tId]) {
            continue;
          }

          var tNaam = toetsNamen[tId] || rows[i][1] || "Onbekende Toets";
          
          if (!perToets[tNaam]) {
            perToets[tNaam] = baseStats();
          }
          
          var status = String(rows[i][6]).toLowerCase();
          
          all.totaal++; 
          perToets[tNaam].totaal++;
          
          if (status === "gestart" || status === "bezig") { 
            all.bezig++; 
            perToets[tNaam].bezig++; 
          } else if (status.indexOf("duplicaat") === -1) { 
            all.verwerkt++; 
            perToets[tNaam].verwerkt++; 
            
            if (status.indexOf("geslaagd") > -1) { 
              all.geslaagd++; 
              perToets[tNaam].geslaagd++; 
            } else if (status.indexOf("gezakt") > -1) { 
              all.gezakt++; 
              perToets[tNaam].gezakt++; 
            }
          }
        }
      }
    }
    
    var shSes = ss.getSheetByName(TAB_SESSIES);
    if (shSes) {
      var sessies = shSes.getDataRange().getValues();
      for (var k = 1; k < sessies.length; k++) {
        if (sessies[k][10] === "OPEN") {
          var tIdSes = sessies[k][1];
          if (!allowedIds._isAdmin && !allowedIds[tIdSes]) {
            continue;
          }
          
          var tNaamSes = toetsNamen[tIdSes] || "Onbekende Toets";
          if (!perToets[tNaamSes]) {
            perToets[tNaamSes] = baseStats();
          }
          
          all.bezig++; 
          perToets[tNaamSes].bezig++;
        }
      }
    }
    
    all.totaal = all.verwerkt + all.bezig;
    
    for (var key in perToets) {
      perToets[key].totaal = perToets[key].verwerkt + perToets[key].bezig;
    }
    
    return { all: all, perToets: perToets };
  } catch(e) { 
    return { all: { totaal: 0, bezig: 0, verwerkt: 0, geslaagd: 0, gezakt: 0 }, perToets: {} }; 
  }
}

function getToetsenOverzicht(clientEmail) {
  try {
    var emailToUse = getViewAsEmail_(clientEmail);
    var allowedIds = getAllowedToetsIds_(emailToUse);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shIdx = ss.getSheetByName(TAB_INDEX);
    var shDb = ss.getSheetByName(TAB_DB);
    
    // Zorg automatisch dat de sheet de juiste kolomkoppen heeft voor video
    if (shIdx) {
      var headerM = shIdx.getRange(1, 13).getValue();
      if (!headerM) {
        shIdx.getRange(1, 13, 1, 2).setValues([["VideoUrl", "VideoDuur"]]);
      }
    }
    
    var pendingCount = 0;
    if (allowedIds._isAdmin) {
       var shAanvraag = ss.getSheetByName(TAB_AANVRAAG_INDEX);
       if (shAanvraag) {
          pendingCount = Math.max(0, shAanvraag.getLastRow() - 1);
       }
    }
    
    if(!shIdx) {
      return { tests: [], canCreate: false, isAdmin: false, pendingCount: pendingCount, userEmail: emailToUse };
    }
    
    var rows = shIdx.getDataRange().getValues();
    if (rows.length < 2) {
      return { tests: [], canCreate: false, isAdmin: false, pendingCount: pendingCount, userEmail: emailToUse };
    }
    
    var dbData = shDb ? shDb.getDataRange().getValues() : [];
    var output = [];
    
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) {
        continue;
      }
      
      if (!allowedIds._isAdmin && !allowedIds[r[0]]) {
        continue;
      }
      
      var mandCount = dbData.filter(function(v) { 
        return v[0] == r[0] && v[8] === true; 
      }).length;
      
      output.push({ 
        id: r[0], 
        naam: r[1], 
        managerName: r[5], 
        managerEmail: r[6], 
        norm: parseNorm(r[7]), 
        vragen_totaal: r[8], 
        vragen_verplicht: mandCount,
        hasVideo: !!(r[12] && String(r[12]).trim() !== ""),
        videoUrl: r[12] ? String(r[12]).trim() : "",
        videoDuur: parseInt(r[13], 10) || 0
      });
    }
    
    return { tests: output.reverse(), canCreate: allowedIds.canCreate, isAdmin: allowedIds._isAdmin, pendingCount: pendingCount, userEmail: emailToUse };
  } catch(e) {
    return { tests: [], canCreate: false, isAdmin: false, pendingCount: 0, userEmail: "" };
  }
}

function getQuizForEdit(toetsId) {
  try {
    var email = "";
    try { email = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(e) {}
    
    var allowedIds = getAllowedToetsIds_(email);
    if (!allowedIds._isAdmin && !allowedIds[toetsId]) {
      throw new Error("Geen rechten om deze toets te bewerken.");
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shIdx = ss.getSheetByName(TAB_INDEX);
    var shDb = ss.getSheetByName(TAB_DB);
    
    // Zorg automatisch dat de sheet de juiste kolomkoppen heeft voor video
    if (shIdx) {
      var headerM = shIdx.getRange(1, 13).getValue();
      if (!headerM) {
        shIdx.getRange(1, 13, 1, 2).setValues([["VideoUrl", "VideoDuur"]]);
      }
    }
    
    if(!shIdx || !shDb) {
      throw new Error("Database tabbladen ontbreken.");
    }
    
    var idxRows = shIdx.getDataRange().getValues();
    var meta = null;
    
    for (var i = 1; i < idxRows.length; i++) {
      if (idxRows[i][0] == toetsId) {
        meta = { 
          id: idxRows[i][0], 
          naam: idxRows[i][1], 
          desc: idxRows[i][2], 
          regio: idxRows[i][3] || "", 
          klantnaam: idxRows[i][4] || "", 
          manName: idxRows[i][5], 
          manEmail: idxRows[i][6], 
          norm: parseNorm(idxRows[i][7]),
          stuurManagerMail: (idxRows[i][10] !== false && String(idxRows[i][10]).toUpperCase() !== "FALSE"),
          toonAntwoorden: (idxRows[i][11] !== false && String(idxRows[i][11]).toUpperCase() !== "FALSE"),
          videoUrl: idxRows[i][12] ? String(idxRows[i][12]).trim() : "",
          videoDuur: parseInt(idxRows[i][13], 10) || 0,
          rowIndex: i+1 
        };
        break;
      }
    }
    
    if (!meta) {
      throw new Error("Toets niet gevonden.");
    }
    
    var db = shDb.getDataRange().getValues();
    var dbDisplay = shDb.getDataRange().getDisplayValues();
    var vragen = [];
    
    for (var r = 1; r < db.length; r++) {
      if (db[r][0] == toetsId) {
        vragen.push({ 
          vraag: String(dbDisplay[r][1] || ""), 
          img: String(db[r][2] || ""), 
          o1: String(dbDisplay[r][3] || ""), 
          o2: String(dbDisplay[r][4] || ""), 
          o3: String(dbDisplay[r][5] || ""), 
          o4: String(dbDisplay[r][6] || ""), 
          correct: normalizeCorrectIdx_(db[r][7]), 
          verplicht: db[r][8] === true, 
          uitleg: String(dbDisplay[r][9] || "") 
        });
      }
    }
    
    return { meta: meta, vragen: vragen };
  } catch(e) {
    throw new Error("Fout bij ophalen toets data: " + e.message);
  }
}

function saveQuizChanges(data) {
  try {
    var allowedIds = getAllowedToetsIds_();
    if (data.meta.id) {
      if (!allowedIds._isAdmin && !allowedIds[data.meta.id]) {
        throw new Error("Geen rechten om deze toets te wijzigen.");
      }
    } else if (!allowedIds.canCreate) {
      throw new Error("Geen rechten om een toets aan te maken.");
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shIdx = ss.getSheetByName(TAB_INDEX);
    var shDb = ss.getSheetByName(TAB_DB);
    
    // Zorg automatisch dat de sheet de juiste kolomkoppen heeft voor video
    if (shIdx) {
      var headerM = shIdx.getRange(1, 13).getValue();
      if (!headerM) {
        shIdx.getRange(1, 13, 1, 2).setValues([["VideoUrl", "VideoDuur"]]);
      }
    }
    
    if (!data.meta.id) {
      var newId = Utilities.getUuid().slice(0, 8);
      data.meta.id = newId;
      
      shIdx.insertRowBefore(2);
      shIdx.getRange(2, 1, 1, 14).setValues([[
        newId, 
        data.meta.naam, 
        data.meta.desc || "", 
        data.meta.regio || "", 
        data.meta.klantnaam || "", 
        data.meta.manName || "", 
        data.meta.manEmail || "", 
        data.meta.norm, 
        data.vragen.length, 
        new Date(),
        data.meta.stuurManagerMail,
        data.meta.toonAntwoorden !== false,
        data.meta.videoUrl || "",
        parseInt(data.meta.videoDuur, 10) || 0
      ]]);
      
      writeLog_("TOETS_AANGEMAAKT", "Naam: " + data.meta.naam);
    } else {
      var indexRows = shIdx.getDataRange().getValues();
      var r = -1;
      for (var row = 1; row < indexRows.length; row++) {
        if (String(indexRows[row][0]) === String(data.meta.id)) {
          r = row + 1;
          break;
        }
      }
      if (r < 0) {
        throw new Error("Toets niet gevonden.");
      }
      
      shIdx.getRange(r, 2, 1, 8).setValues([[
        data.meta.naam, 
        data.meta.desc || "", 
        data.meta.regio || "", 
        data.meta.klantnaam || "", 
        data.meta.manName || "", 
        data.meta.manEmail || "", 
        data.meta.norm, 
        data.vragen.length
      ]]);
      
      shIdx.getRange(r, 11).setValue(data.meta.stuurManagerMail);
      shIdx.getRange(r, 12).setValue(data.meta.toonAntwoorden !== false);
      shIdx.getRange(r, 13).setValue(data.meta.videoUrl || "");
      shIdx.getRange(r, 14).setValue(parseInt(data.meta.videoDuur, 10) || 0);
      
      var dbData = shDb.getDataRange().getValues();
      for (var i = dbData.length - 1; i >= 1; i--) { 
        if (dbData[i][0] == data.meta.id) {
          shDb.deleteRow(i + 1); 
        }
      }
      
      writeLog_("TOETS_GEWIJZIGD", "Naam: " + data.meta.naam);
    }
    
    if (data.vragen.length > 0) {
      var newRows = data.vragen.map(function(q) { 
        return [
          data.meta.id, 
          q.vraag, 
          q.img, 
          q.o1, 
          q.o2, 
          q.o3, 
          q.o4, 
          q.correct, 
          q.verplicht, 
          q.uitleg || ""
        ]; 
      });
      shDb.insertRowsBefore(2, newRows.length);
      shDb.getRange(2, 1, newRows.length, 10).setValues(newRows);
    }
    
    clearSystemCaches(); 
    return { success: true };
  } catch (e) { 
    throw new Error("Opslaan mislukt: " + e.message); 
  }
}

function importeerToetsInDatabase(url) {
  try {
    if (!getAllowedToetsIds_()._isAdmin) {
      throw new Error("Alleen een beheerder mag toetsen importeren.");
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var src = SpreadsheetApp.openByUrl(url).getSheets()[0];
    
    if (src.getLastRow() < 8) {
       throw new Error("Validatie mislukt: Er zijn geen vragen gevonden in deze sheet (vanaf rij 8). Gebruik het officiële G4S Aanleverformulier.");
    }
    
    var meta = src.getRange(4, 1, 2, 11).getValues();
    var toetsNaam = meta[0][2];
    
    if (!toetsNaam || String(toetsNaam).trim() === "") {
       throw new Error("Validatie mislukt: De toetsnaam ontbreekt op de verwachte locatie (Cel C4). Controleer of dit het officiële G4S Aanleverformulier is.");
    }
    
    var id = Utilities.getUuid().slice(0, 8);
    var norm = parseNorm(meta[1][6]);
    
    var idxRow = [
      id, 
      toetsNaam,  
      meta[1][2], 
      meta[0][4], 
      meta[1][4], 
      meta[0][8], 
      meta[1][8], 
      norm,        
      meta[0][6], 
      new Date(),
      true,
      true
    ];
    
    var shIdx = ss.getSheetByName(TAB_INDEX);
    shIdx.insertRowBefore(2);
    shIdx.getRange(2, 1, 1, 12).setValues([idxRow]);
    
    var vr = src.getRange(8, 1, src.getLastRow() - 7, 12).getValues();
    var vrDisplay = src.getRange(8, 1, src.getLastRow() - 7, 12).getDisplayValues();
    var dbRows = [];
    
    for (var i = 0; i < vr.length; i++) {
      if (!vr[i][2]) {
        break; 
      }
      
      dbRows.push([
        id, 
        String(vrDisplay[i][2]).trim(),  
        String(vrDisplay[i][3]).trim(),  
        vrDisplay[i][4],                 
        vrDisplay[i][5],                 
        vrDisplay[i][6],                 
        vrDisplay[i][7],                 
        normalizeCorrectIdx_(vr[i][8]),                     
        (String(vr[i][9]).toLowerCase().indexOf("ja") > -1), 
        vrDisplay[i][10]                 
      ]);
    }
    
    if (dbRows.length > 0) {
      var shDb = ss.getSheetByName(TAB_DB);
      shDb.insertRowsBefore(2, dbRows.length);
      shDb.getRange(2, 1, dbRows.length, 10).setValues(dbRows);
    }
    
    writeLog_("TOETS_GEIMPORTEERD", "Naam: " + toetsNaam);
    clearSystemCaches(); 
    return { success: true, message: "Geïmporteerd: " + dbRows.length + " vragen." };
  } catch(e) { 
    throw new Error("Importeren mislukt: " + e.message); 
  }
}

// ----------------------------------------------------------------------------------------------------
// EMAIL & UITNODIGINGEN 
// ----------------------------------------------------------------------------------------------------

function genereerEnVerstuur(cfg) {
  try {
    var allowedIds = getAllowedToetsIds_();
    if (!allowedIds._isAdmin && !allowedIds[cfg.toetsId]) {
      throw new Error("Geen rechten om deze toets te versturen.");
    }

    var employeeMap = getEmployeeMapping_();
    
    var rawEmails = (cfg.emails||"").split(/[\n,;]+/); 
    var validInputs = [];
    var invalidEmails = [];
    
    var accTypeStr = String(cfg.accType || "Persoonlijk");
    var isPersoonlijk = accTypeStr.indexOf("Persoonlijk") > -1;
    var isBinnenG4S = accTypeStr.indexOf("Binnen G4S") > -1;
    var isGedeeld = accTypeStr.indexOf("Gedeeld") > -1;

    for (var e = 0; e < rawEmails.length; e++) {
      var t = rawEmails[e].trim(); 
      if (t === "") continue;
      
      var tLower = t.toLowerCase();
      if (tLower.indexOf("@") < 0) {
        invalidEmails.push(t); 
        continue;
      }
      
      if (isPersoonlijk) {
        var validRecordFound = false;
        var targetMail = "";
        var targetUpn = "";
        
        if (employeeMap[tLower] && employeeMap[tLower].upn === tLower && employeeMap[tLower].source === "HR") {
            validRecordFound = true;
            targetUpn = tLower;
            var empObj = employeeMap[tLower];
            targetMail = tLower;
            if (!empObj.isGroepsleider && empObj.privateEmail && empObj.privateEmail.indexOf("@") > -1) {
                targetMail = empObj.privateEmail;
            }
        } 
        else {
            for (var key in employeeMap) {
                if (employeeMap[key].privateEmail === tLower && employeeMap[key].source === "HR") {
                    validRecordFound = true;
                    targetUpn = employeeMap[key].upn;
                    targetMail = tLower; 
                    break;
                }
            }
        }

        if (!validRecordFound) {
          invalidEmails.push(t); 
        } else {
          var alToegevoegd = false;
          for(var k=0; k<validInputs.length; k++) {
            if (validInputs[k].upn === targetUpn) alToegevoegd = true;
          }
          
          if (!alToegevoegd) {
            validInputs.push({
              upn: targetUpn,
              sendTo: targetMail
            });
          }
        }
      } else if (isBinnenG4S) {
        if (tLower.indexOf("@nl.g4s.com") === -1 && tLower.indexOf("@g4s.com") === -1) {
          invalidEmails.push(t);
        } else {
          var alToegevoegd = false;
          for(var k=0; k<validInputs.length; k++) {
            if (validInputs[k].upn === tLower) alToegevoegd = true;
          }
          if (!alToegevoegd) {
            validInputs.push({
              upn: tLower,
              sendTo: tLower
            });
          }
        }
      } else if (isGedeeld) {
        if (!employeeMap[tLower] || employeeMap[tLower].source !== "SHARED") {
           invalidEmails.push(t);
        } else {
            var alToegevoegd = false;
            for(var k=0; k<validInputs.length; k++) {
              if (validInputs[k].upn === tLower) alToegevoegd = true;
            }

            if (!alToegevoegd) {
              validInputs.push({
                upn: tLower,
                sendTo: tLower
              });
            }
        }
      } else {
        invalidEmails.push(t);
      }
    }
    
    if (invalidEmails.length > 0 && validInputs.length === 0) {
      return { success: false, isValidationError: true, invalidEmails: invalidEmails };
    }
    
    if (validInputs.length === 0) {
      return { success: false, message: "Geen geldige e-mailadressen ingevoerd." };
    }
    
    var scriptUrl = ScriptApp.getService().getUrl();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shSes = ss.getSheetByName(TAB_SESSIES);
    
    if (!shSes) {
      shSes = ss.insertSheet(TAB_SESSIES);
      shSes.appendRow(["Type", "ToetsID", "Norm", "CC", "Manager", "Datum", "Tijd", "Herhaling", "Email", "SID", "Status", "GebruiktOp", "AantalVragen", "Verstuurd Naar"]);
    }
    
    var shIdx = ss.getSheetByName(TAB_INDEX);
    if (!shIdx) {
      throw new Error("Database index niet gevonden.");
    }
    
    var idxRows = shIdx.getDataRange().getValues();
    var meta = null;
    
    for (var i = 1; i < idxRows.length; i++) {
      if (idxRows[i][0] == cfg.toetsId) { 
        meta = idxRows[i]; 
        break; 
      }
    }
    
    if(!meta) {
      return {success:false, message:"Toets niet gevonden in database."};
    }
    
    var rawManNaam = meta[5] || "G4S Operation Support Team"; 
    var managerNaam = String(rawManNaam).split(',')[0].trim();
    var managerEmail = meta[6] || ""; 
    var toetsNaam = meta[1]; 
    var aantalVerzonden = 0;
    
    var hasVideo = meta[12] && String(meta[12]).trim() !== "";
    var vDuurSec = parseInt(meta[13], 10) || 0;
    var vDuurTekst = vDuurSec > 0 ? " (speelduur ca. " + Math.ceil(vDuurSec / 60) + " min.)" : "";
    var videoNoticeHtml = hasVideo ? `
      <div style="background-color: #fff8e1; border-left: 4px solid #f59e0b; padding: 12px 15px; margin: 15px 0; border-radius: 4px; border: 1px solid #fde68a;">
        <p style="margin: 0; color: #b45309; font-size: 14px; line-height: 1.5;">
          <strong>Let op:</strong> Deze toets bevat een verplichte <strong>instructievideo</strong>${vDuurTekst}. Deze video bekijk je direct op het startscherm voordat de vragen beginnen. Zorg dat je geluid aan staat!
        </p>
      </div>
    ` : "";
    var videoRowHtml = hasVideo ? `<p style="margin: 0; color: #333;"><strong>Instructievideo:</strong> Ja${vDuurTekst}</p>` : "";

    var logoUrl = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : "https://www.g4s.com/-/media/g4s/global/images/logos/g4s---an-allied-universal-company-logo/g4s_one-company_logo.png";

    for (var v = 0; v < validInputs.length; v++) {
      var validEm = validInputs[v].upn; 
      var sendToMail = validInputs[v].sendTo;
      
      var sid = Utilities.getUuid().replace(/-/g,"").slice(0,14);
      
      var safeScriptUrl = scriptUrl;
      var savedUrl = PropertiesService.getScriptProperties().getProperty('SAVED_WEBAPP_URL');
      if (savedUrl && safeScriptUrl.indexOf("/dev") > -1) { safeScriptUrl = savedUrl; }
      
      safeScriptUrl = safeScriptUrl.replace(/\/dev$/i, "/exec");
      
      var baseAppUrl = safeScriptUrl + "?toetsId=" + encodeURIComponent(cfg.toetsId) + "&sid=" + sid;
      var link = "https://accounts.google.com/AccountChooser?hl=nl&continue=" + encodeURIComponent(baseAppUrl) + "&hd=nl.g4s.com&prompt=select_account";
      
      shSes.appendRow([
        "APP_LINK", 
        cfg.toetsId, 
        cfg.minScore + "%", 
        cfg.ccEmails, 
        managerEmail, 
        new Date(), 
        cfg.maxTijd, 
        cfg.herhaling, 
        validEm, 
        sid, 
        "OPEN", 
        "", 
        cfg.aantalVragen,
        sendToMail
      ]);
      
      var emailHtml = `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f6f9; padding: 20px; border-radius: 8px;">
          <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; border-top: 4px solid #D32F2F; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px;">
              <tr>
                <td style="vertical-align: middle;"><h2 style="color: #232323; font-size: 20px; margin: 0;">Uitnodiging Toets</h2></td>
                <td style="text-align: right; vertical-align: middle;"><img src="${logoUrl}" alt="G4S Logo" style="height: 50px;"></td>
              </tr>
            </table>
            <p style="color: #333; font-size: 15px; line-height: 1.6;">Beste collega,</p>
            <p style="color: #333; font-size: 15px; line-height: 1.6;">Je bent uitgenodigd om een kennistoets te maken. Deze uitnodiging en de inhoud van de toets zijn bepaald in opdracht van jouw Operationeel Manager: <strong>${managerNaam}</strong>.</p>
            
            <div style="background-color: #e3f2fd; border-left: 4px solid #1565c0; padding: 15px; margin: 20px 0; border: 1px solid #cbe3f5;">
              <p style="margin: 0; color: #0d47a1; font-size: 14px; line-height: 1.5;"><strong>Belangrijk:</strong> Je ontvangt deze notificatie op dit e-mailadres voor de communicatie. Om de toets daadwerkelijk te kunnen maken, zal er gevraagd worden om in te loggen in de beveiligde omgeving. Gebruik hiervoor jouw <strong>G4S-inloggegevens</strong> (je G4SUPN en bijbehorend wachtwoord).</p>
            </div>

            <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0; border: 1px solid #eee;">
              <p style="margin: 0 0 8px 0; color: #333;"><strong>Toets:</strong> ${toetsNaam}</p>
              <p style="margin: 0 0 8px 0; color: #333;"><strong>Aantal vragen:</strong> ${cfg.aantalVragen}</p>
              <p style="margin: 0 ${hasVideo ? '0 8px 0' : ''}; color: #333;"><strong>Norm:</strong> ${cfg.minScore}%</p>
              ${videoRowHtml}
            </div>
            ${videoNoticeHtml}
            <div style="text-align: center; margin: 30px 0;">
              <a href="${link}" target="_blank" rel="noopener noreferrer" style="background-color: #D32F2F; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; text-transform: uppercase;">START TOETS IN PORTAL</a>
            </div>
            <p style="color: #666; font-size: 13px; line-height: 1.5; margin-top: 30px;">
              Met vriendelijke groet,<br><strong>G4S Operation Support Team</strong>
            </p>
          </div>
        </div>
      `;
        
      var subjectLine = "Uitnodiging Toets: " + toetsNaam;
      if (isGedeeld && validInputs.length > 1) {
        subjectLine += " (Link " + (v + 1) + " van " + validInputs.length + ")";
      }
        
      try {
        MailApp.sendEmail({ to: sendToMail, subject: subjectLine, htmlBody: emailHtml });
        aantalVerzonden++;
      } catch (err) {
        Logger.log("Fout bij verzenden naar " + sendToMail + ": " + err.message);
      }
    }
    
    writeLog_("UITNODIGINGEN_VERSTUURD", "Toets: " + toetsNaam + " (" + aantalVerzonden + " personen)");
    clearSystemCaches(); 
    
    return { 
      success: true, 
      message: "Succesvol " + aantalVerzonden + " uitnodiging(en) verstuurd!",
      isValidationError: (invalidEmails.length > 0),
      invalidEmails: invalidEmails
    };
  } catch(e) { 
    throw new Error("Versturen mislukt: " + e.message); 
  }
}

// ----------------------------------------------------------------------------------------------------
// HERKANSINGEN
// ----------------------------------------------------------------------------------------------------

function stuurHerkansing(req) {
  try {
    var allowedIds = getAllowedToetsIds_();
    if (!allowedIds._isAdmin && !allowedIds[req.toetsId]) {
      throw new Error("Geen rechten om deze herkansing te versturen.");
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shSes = ss.getSheetByName(TAB_SESSIES);
    var shIdx = ss.getSheetByName(TAB_INDEX);
    var shRes = ss.getSheetByName(TAB_RESULTATEN);
    
    var idxRows = shIdx.getDataRange().getValues();
    var toetsNaam = req.toetsNaam;
    var norm = 70;
    var managerEmail = "";
    var hasVideo = false;
    var vDuurSec = 0;
    for (var i = 1; i < idxRows.length; i++) {
      if (idxRows[i][0] == req.toetsId) {
        toetsNaam = idxRows[i][1];
        norm = parseNorm(idxRows[i][7]);
        managerEmail = idxRows[i][6];
        hasVideo = idxRows[i][12] && String(idxRows[i][12]).trim() !== "";
        vDuurSec = parseInt(idxRows[i][13], 10) || 0;
        break;
      }
    }

    var ccEmails = "";
    var maxTijd = 0;
    var aantalVragen = 0;
    var herhaling = 0;
    
    if (req.sid) {
      var sRows = shSes.getDataRange().getValues();
      for (var j = sRows.length - 1; j >= 1; j--) {
        if (sRows[j][9] === req.sid) {
          ccEmails = sRows[j][3];
          maxTijd = sRows[j][6];
          herhaling = sRows[j][7];
          aantalVragen = sRows[j][12];
          break;
        }
      }
    }

    var dataRes = shRes.getDataRange().getValues();
    for (var k = dataRes.length - 1; k >= 1; k--) {
      var rowEmail = String(dataRes[k][0]).toLowerCase();
      if (rowEmail.indexOf(req.email.toLowerCase()) > -1 && dataRes[k][7] == req.toetsId && dataRes[k][9] === req.sid) {
        var oudeStatus = dataRes[k][6];
        if (oudeStatus.indexOf("Herkansing toegewezen") === -1) {
          shRes.getRange(k + 1, 7).setValue(oudeStatus + " (Herkansing toegewezen)");
        }
        break;
      }
    }

    var targetMail = req.email; 
    var targetUpn = req.email;
    try {
      var employeeMap = getEmployeeMapping_();
      var searchEmail = req.email.toLowerCase().trim();
      if (employeeMap[searchEmail]) {
        var empObj = employeeMap[searchEmail];
        if (!empObj.isGroepsleider && empObj.privateEmail && empObj.privateEmail.indexOf("@") > -1) {
          targetMail = empObj.privateEmail;
        }
      }
    } catch(e) {
      Logger.log("Fout bij ophalen prive adres voor herkansing: " + e.message);
    }

    var newSid = Utilities.getUuid().replace(/-/g,"").slice(0,14);
    var scriptUrl = ScriptApp.getService().getUrl();
    var savedUrl = PropertiesService.getScriptProperties().getProperty('SAVED_WEBAPP_URL');
    if (savedUrl && scriptUrl.indexOf("/dev") > -1) { scriptUrl = savedUrl; }
    scriptUrl = scriptUrl.replace(/\/dev$/i, "/exec");
    
    var baseAppUrl = scriptUrl + "?toetsId=" + encodeURIComponent(req.toetsId) + "&sid=" + newSid;
    var link = "https://accounts.google.com/AccountChooser?hl=nl&continue=" + encodeURIComponent(baseAppUrl) + "&hd=nl.g4s.com&prompt=select_account";

    shSes.appendRow([
      "APP_LINK", 
      req.toetsId, 
      norm + "%", 
      ccEmails, 
      managerEmail, 
      new Date(), 
      maxTijd, 
      herhaling, 
      targetUpn, 
      newSid, 
      "OPEN", 
      "", 
      aantalVragen,
      targetMail 
    ]);

    var logoUrl = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : "https://www.g4s.com/-/media/g4s/global/images/logos/g4s---an-allied-universal-company-logo/g4s_one-company_logo.png";
    var emailHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f6f9; padding: 20px; border-radius: 8px;">
        <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; border-top: 4px solid #f59e0b; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px;">
            <tr>
              <td style="vertical-align: middle;"><h2 style="color: #232323; font-size: 20px; margin: 0;">Uitnodiging Herkansing Toets</h2></td>
              <td style="text-align: right; vertical-align: middle;"><img src="${logoUrl}" alt="G4S Logo" style="height: 50px;"></td>
            </tr>
          </table>
          <p style="color: #333; font-size: 15px; line-height: 1.6;">Beste collega,</p>
          <p style="color: #333; font-size: 15px; line-height: 1.6;">Je manager heeft je een <strong>herkansing</strong> toegekend voor een kennistoets die je eerder niet succesvol hebt afgerond.</p>
          
          <div style="background-color: #fef3c7; border-left: 4px solid #f59e0b; padding: 15px; margin: 20px 0; border-radius: 4px; border-top: 1px solid #fde68a; border-right: 1px solid #fde68a; border-bottom: 1px solid #fde68a;">
            <p style="margin: 0; color: #b45309; font-size: 14px; line-height: 1.5;"><strong>Belangrijk:</strong> Log in met je <strong>G4S-inloggegevens</strong> (je G4SUPN en bijbehorend wachtwoord) om de toets opnieuw te maken.</p>
          </div>

          <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0; border: 1px solid #eee;">
            <p style="margin: 0 0 8px 0; color: #333;"><strong>Toets:</strong> ${toetsNaam}</p>
            <p style="margin: 0 ${hasVideo ? '0 8px 0' : ''}; color: #333;"><strong>Norm:</strong> ${norm}%</p>
            ${hasVideo ? `<p style="margin: 0; color: #333;"><strong>Instructievideo:</strong> Ja${vDuurSec > 0 ? ' (ca. ' + Math.ceil(vDuurSec / 60) + ' min.)' : ''}</p>` : ''}
          </div>
          ${hasVideo ? `
            <div style="background-color: #fff8e1; border-left: 4px solid #f59e0b; padding: 12px 15px; margin: 15px 0; border-radius: 4px; border: 1px solid #fde68a;">
              <p style="margin: 0; color: #b45309; font-size: 14px; line-height: 1.5;">
                <strong>Let op:</strong> Deze toets bevat een verplichte <strong>instructievideo</strong>. Deze video bekijk je direct op het startscherm voordat de vragen beginnen. Zorg dat je geluid aan staat!
              </p>
            </div>
          ` : ''}
          <div style="text-align: center; margin: 30px 0;">
            <a href="${link}" target="_blank" rel="noopener noreferrer" style="background-color: #f59e0b; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; text-transform: uppercase;">START HERKANSING</a>
          </div>
          <p style="color: #666; font-size: 13px; line-height: 1.5; margin-top: 30px;">
            Succes met deze poging!<br><strong>G4S Operation Support Team</strong>
          </p>
        </div>
      </div>
    `;
    
    try {
      MailApp.sendEmail({ to: targetMail, subject: "Herkansing Toets: " + toetsNaam, htmlBody: emailHtml });
    } catch (err) {
      Logger.log("Fout bij verzenden herkansing naar " + targetMail + ": " + err.message);
    }
    
    writeLog_("HERKANSING_VERSTUURD", "Naar: " + targetMail + " (Kandidaat: " + targetUpn + ") - Toets: " + toetsNaam);
    clearSystemCaches();
    
    return { success: true, message: "Herkansing succesvol verzonden naar " + targetMail };
  } catch(e) {
    throw new Error("Fout bij aanmaken herkansing: " + e.message);
  }
}

// ----------------------------------------------------------------------------------------------------
// OVERIGE FUNCTIES (POWERBI, HERINNERINGEN, CACHE ETC.)
// ----------------------------------------------------------------------------------------------------

function getPowerBIData(clientEmail, forceRefresh) {
  if (forceRefresh) {
     clearSystemCaches();
  }
  var email = getViewAsEmail_(clientEmail);
  var version = getGlobalCacheVersion();
  
  var cacheKey = 'pbi_' + Utilities.base64Encode(email).substring(0,25) + '_' + version;
  
  var cached = getChunkedCache(cacheKey);
  if (cached) return cached;
  
  var data = getPowerBIData_calc_(email);
  putChunkedCache(cacheKey, data, 21600); 
  return data;
}

function getPowerBIData_calc_(email) {
  try {
    var allowedIds = getAllowedToetsIds_(email);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    function formatTimeHHMMSS(mins) {
      var mFloat = parseFloat(String(mins).replace(',', '.')); 
      if (isNaN(mFloat) || mFloat === 0) {
        return "00:00:00";
      }
      var totalSeconds = Math.round(mFloat * 60);
      var h = Math.floor(totalSeconds / 3600); 
      var m = Math.floor((totalSeconds % 3600) / 60); 
      var s = totalSeconds % 60;
      return (h < 10 ? "0" + h : h) + ":" + (m < 10 ? "0" + m : m) + ":" + (s < 10 ? "0" + s : s);
    }
    
    var shIdx = ss.getSheetByName(TAB_INDEX); 
    var toetsMap = {};
    if (shIdx) {
      var idxData = shIdx.getDataRange().getValues();
      for(var i = 1; i < idxData.length; i++) {
        var nRaw = parseFloat(String(idxData[i][7]).replace('%','').replace(',','.'));
        toetsMap[idxData[i][0]] = { 
          naam: idxData[i][1], 
          regio: idxData[i][3] || "Onbekend", 
          locatie: idxData[i][4] || "Onbekend", 
          manager: idxData[i][5] || "Onbekend", 
          norm: !isNaN(nRaw) ? nRaw.toFixed(2).replace('.', ',') + "%" : "-" 
        };
      }
    }
    
    var shEmp = ss.getSheetByName(TAB_EMP); 
    var empMap = {};
    if (shEmp) {
      var empData = shEmp.getDataRange().getValues(); 
      if(empData.length > 0) {
        var headers = empData[0]; 
        var emailIdx = headers.indexOf("EMAIL"); 
        var fnIdx = headers.indexOf("firstname"); 
        var lnIdx = headers.indexOf("lastname");
        
        if(emailIdx > -1) { 
          for(var j = 1; j < empData.length; j++) {
            var em = String(empData[j][emailIdx]).toLowerCase().trim(); 
            if(!em) {
              continue;
            }
            var fullName = [empData[j][fnIdx], empData[j][lnIdx]].filter(Boolean).join(" ");
            empMap[em] = { naam: fullName || em };
          }
        }
      }
    }
    
    var shSes = ss.getSheetByName(TAB_SESSIES); 
    var sesData = shSes ? shSes.getDataRange().getValues() : [];
    var sesMap = {}; 
    for(var m = 1; m < sesData.length; m++) { 
      if(sesData[m][9]) {
        sesMap[sesData[m][9]] = { time: sesData[m][6], cc: String(sesData[m][3]||"Geen") }; 
      }
    }
    
    var flatData = []; 
    var processedSessions = {}; 
    var shRes = ss.getSheetByName(TAB_RESULTATEN);
    
    if (shRes) {
      var resData = shRes.getDataRange().getValues();
      for(var k = 1; k < resData.length; k++) {
        
        var rawStatus = String(resData[k][6]).toLowerCase();
        if (rawStatus.indexOf("duplicaat") > -1) {
          continue; 
        }
        
        var tId = resData[k][7]; 
        if (!allowedIds._isAdmin && !allowedIds[tId]) {
          continue;
        }
        
        var sCode = resData[k][9]; 
        if(sCode) {
          processedSessions[sCode] = true;
        }
        
        var tInfo = toetsMap[tId] || {regio: "Onbekend", locatie: "Onbekend", manager: "Onbekend", norm: "-"};
        var eInfo = empMap[String(resData[k][0]).toLowerCase()] || {naam: resData[k][0]};
        
        var scRaw = resData[k][5]; 
        var scPerc = "-";
        
        if (scRaw) {
          var pctMatch = String(scRaw).match(/(\d+)\s*%/);
          if (pctMatch) {
            scPerc = parseFloat(pctMatch[1]).toFixed(2).replace('.', ',') + "%";
          }
        }
        
        var sInfo = sesMap[sCode] || {time: 0, cc: "Geen"};
        
        var dGemaakt = resData[k][3];
        var dGemaaktStr = "-";
        if (dGemaakt instanceof Date) {
          dGemaaktStr = dGemaakt.toLocaleDateString('nl-NL');
        } else if (dGemaakt) {
          dGemaaktStr = String(dGemaakt).split(' ')[0];
        }

        var tInst = sInfo.time;
        var tDeel = resData[k][4];
        var tijdInstellingStr = (!tInst || tInst == 0) ? "Geen limiet" : formatTimeHHMMSS(tInst);
        var tijdDeelnemerStr = (!tDeel && tDeel !== 0) ? "-" : formatTimeHHMMSS(tDeel);
        
        var outStatus = "Gezakt";
        if (rawStatus.indexOf("geslaagd") > -1) {
            outStatus = "Geslaagd";
        } else if (rawStatus.indexOf("tijd overschreden") > -1) {
            outStatus = "Gezakt (Tijd overschreden)";
        }
        
        flatData.push({ 
          datumGemaakt: String(dGemaaktStr || "-"), 
          email: String(resData[k][0] || "-"), 
          naam: String(eInfo.naam || "-"), 
          ccLijst: String(sInfo.cc || "Geen"), 
          regio: String(tInfo.regio || "Onbekend"), 
          locatie: String(tInfo.locatie || "Onbekend"), 
          manager: String(tInfo.manager || "Onbekend"), 
          status: String(outStatus), 
          score: String(scPerc), 
          norm: String(tInfo.norm), 
          tijdInstelling: String(tijdInstellingStr), 
          tijdDeelnemer: String(tijdDeelnemerStr),
          details: String(resData[k][8] || "[]") 
        });
      }
    }
    
    for(var l = 1; l < sesData.length; l++) {
      if(sesData[l][10] === "OPEN" && !processedSessions[sesData[l][9]]) {
        var tIdSes = sesData[l][1]; 
        if (!allowedIds._isAdmin && !allowedIds[tIdSes]) {
          continue;
        }
        
        var tInfoSes = toetsMap[tIdSes] || {regio: "Onbekend", locatie: "Onbekend", manager: "Onbekend", norm: "-"};
        var eInfoSes = empMap[String(sesData[l][8]).toLowerCase()] || {naam: sesData[l][8]};
        
        var dAangemaaktSes = sesData[l][5];
        var dAangemaaktSesStr = "-";
        if (dAangemaaktSes instanceof Date) {
          dAangemaaktSesStr = dAangemaaktSes.toLocaleDateString('nl-NL');
        } else if (dAangemaaktSes) {
          dAangemaaktSesStr = String(dAangemaaktSes).split(' ')[0];
        }

        var tInstSes = sesData[l][6];
        var tijdInstSesStr = (!tInstSes || tInstSes == 0) ? "Geen limiet" : formatTimeHHMMSS(tInstSes);

        flatData.push({ 
          datumGemaakt: String(dAangemaaktSesStr || "-"), 
          email: String(sesData[l][8] || "-"), 
          naam: String(eInfoSes.naam || "-"), 
          ccLijst: String(sesData[l][3] || "Geen"), 
          regio: String(tInfoSes.regio || "Onbekend"), 
          locatie: String(tInfoSes.locatie || "Onbekend"), 
          manager: String(tInfoSes.manager || "Onbekend"), 
          status: "Nog te doen", 
          score: "-", 
          norm: String(tInfoSes.norm), 
          tijdInstelling: String(tijdInstSesStr), 
          tijdDeelnemer: "-",
          details: "[]"
        });
      }
    }
    
    return flatData;
  } catch(e) { 
    throw new Error("PowerBI data fout: " + e.message); 
  }
}

function checkEnVerstuurHerinneringen() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hSheet = ss.getSheetByName(TAB_HERINNERINGEN);
    if (!hSheet) return;
    
    var data = hSheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    var now = new Date();
    now.setHours(0,0,0,0);
    var todayStr = now.toISOString().split('T')[0];
    
    var props = PropertiesService.getScriptProperties();
    var lastRun = props.getProperty('lastReminderCheck');
    if (lastRun === todayStr) return; 
    
    var employeeMap = getEmployeeMapping_(); 
    
    var shSes = ss.getSheetByName(TAB_SESSIES);
    var shIdx = ss.getSheetByName(TAB_INDEX);
    var idxRows = shIdx ? shIdx.getDataRange().getValues() : [];
    var scriptUrl = ScriptApp.getService().getUrl();
    var logoUrl = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : "https://www.g4s.com/-/media/g4s/global/images/logos/g4s---an-allied-universal-company-logo/g4s_one-company_logo.png";
    var mutatieGedaan = false;
    
    for (var j = 1; j < data.length; j++) {
      if (data[j][3] === "Gepland") {
        
        var rawDate = data[j][2];
        var planDatum = new Date(rawDate);
        
        if (isNaN(planDatum.getTime()) && typeof rawDate === "string") {
          var p = rawDate.split(' ')[0].split('-');
          if (p.length >= 3) {
            planDatum = new Date(p[2], p[1] - 1, p[0]);
          }
        }
        
        var checkDate = new Date();
        checkDate.setHours(0,0,0,0);
        
        if (!isNaN(planDatum.getTime()) && planDatum.setHours(0,0,0,0) <= checkDate.getTime()) { 
          
          var email = String(data[j][0]).trim().toLowerCase(); 
          
          if (!employeeMap[email]) {
            hSheet.getRange(j+1, 4).setValue("Geannuleerd (Account niet in Dropdown Info)");
            writeLog_("HERHALING_GEANNULEERD", email + " niet meer actief. Toets ID: " + data[j][1]);
            mutatieGedaan = true;
          } else {
            var empObj = employeeMap[email];
            var trueUpn = empObj.upn;
            
            var tid = data[j][1];
            var managerEmail = data[j][4];
            var ccEmails = data[j][5];
            var norm = data[j][6] || 70;
            var aantalVragen = data[j][7] || 10;
            var maxTijd = data[j][8] || 0;
            var herhaling = data[j][9] || 0;
            
            var toetsNaam = "Kennistoets";
            var managerNaam = "G4S Operation Support Team";
            var hasVideoH = false;
            var vDuurSecH = 0;
            for (var k = 1; k < idxRows.length; k++) {
              if (idxRows[k][0] == tid) {
                toetsNaam = idxRows[k][1];
                var rawManNaam = idxRows[k][5] || managerNaam;
                managerNaam = String(rawManNaam).split(',')[0].trim();
                hasVideoH = idxRows[k][12] && String(idxRows[k][12]).trim() !== "";
                vDuurSecH = parseInt(idxRows[k][13], 10) || 0;
                break;
              }
            }
            
            var sid = Utilities.getUuid().replace(/-/g,"").slice(0,14);
            
            var safeScriptUrl = scriptUrl;
            var savedUrl = PropertiesService.getScriptProperties().getProperty('SAVED_WEBAPP_URL');
            if (savedUrl) { safeScriptUrl = savedUrl; } 
            
            safeScriptUrl = safeScriptUrl.replace(/\/dev$/i, "/exec");
            
            var baseAppUrl = safeScriptUrl + "?toetsId=" + encodeURIComponent(tid) + "&sid=" + sid;
            var link = "https://accounts.google.com/AccountChooser?hl=nl&continue=" + encodeURIComponent(baseAppUrl) + "&hd=nl.g4s.com&prompt=select_account";
            
            var sendToMail = email;

            if (shSes) {
              shSes.appendRow(["APP_LINK", tid, norm + "%", ccEmails, managerEmail, new Date(), maxTijd, herhaling, trueUpn, sid, "OPEN", "", aantalVragen, sendToMail]);
              mutatieGedaan = true;
            }
            
            var emailHtml = `
              <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f6f9; padding: 20px; border-radius: 8px;">
                <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; border-top: 4px solid #D32F2F; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px;">
                    <tr>
                      <td style="vertical-align: middle;"><h2 style="color: #232323; font-size: 20px; margin: 0;">Periodieke Herhaling Toets</h2></td>
                      <td style="text-align: right; vertical-align: middle;"><img src="${logoUrl}" alt="G4S Logo" style="height: 50px;"></td>
                    </tr>
                  </table>
                  <p style="color: #333; font-size: 15px; line-height: 1.6;">Beste collega,</p>
                  <p style="color: #333; font-size: 15px; line-height: 1.6;">Het is weer tijd voor je periodieke kennistoets. Deze herinnering en de inhoud van de toets zijn bepaald in opdracht van jouw Operationeel Manager: <strong>${managerNaam}</strong>.</p>
                  
                  <div style="background-color: #e3f2fd; border-left: 4px solid #1565c0; padding: 15px; margin: 20px 0; border-radius: 4px; border-top: 1px solid #cbe3f5; border-right: 1px solid #cbe3f5; border-bottom: 1px solid #cbe3f5;">
                    <p style="margin: 0; color: #0d47a1; font-size: 14px; line-height: 1.5;"><strong>Belangrijk:</strong> Je ontvangt deze notificatie op dit e-mailadres voor de communicatie. Om de toets daadwerkelijk te kunnen maken, zal er gevraagd worden om in te loggen in de beveiligde omgeving. Gebruik hiervoor jouw <strong>G4S-inloggegevens</strong> (je G4SUPN en bijbehorend wachtwoord).</p>
                  </div>

                  <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0; border: 1px solid #eee;">
                    <p style="margin: 0 0 8px 0; color: #333;"><strong>Toets:</strong> ${toetsNaam}</p>
                    <p style="margin: 0 0 8px 0; color: #333;"><strong>Aantal vragen:</strong> ${aantalVragen}</p>
                    <p style="margin: 0 ${hasVideoH ? '0 8px 0' : ''}; color: #333;"><strong>Norm:</strong> ${norm}%</p>
                    ${hasVideoH ? `<p style="margin: 0; color: #333;"><strong>Instructievideo:</strong> Ja${vDuurSecH > 0 ? ' (ca. ' + Math.ceil(vDuurSecH / 60) + ' min.)' : ''}</p>` : ''}
                  </div>
                  ${hasVideoH ? `
                    <div style="background-color: #fff8e1; border-left: 4px solid #f59e0b; padding: 12px 15px; margin: 15px 0; border-radius: 4px; border: 1px solid #fde68a;">
                      <p style="margin: 0; color: #b45309; font-size: 14px; line-height: 1.5;">
                        <strong>Let op:</strong> Deze toets bevat een verplichte <strong>instructievideo</strong>. Deze video bekijk je direct op het startscherm voordat de vragen beginnen. Zorg dat je geluid aan staat!
                      </p>
                    </div>
                  ` : ''}
                  <div style="text-align: center; margin: 30px 0;">
                    <a href="${link}" target="_blank" rel="noopener noreferrer" style="background-color: #D32F2F; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; text-transform: uppercase;">START TOETS IN PORTAL</a>
                  </div>
                  <p style="color: #666; font-size: 13px; line-height: 1.5; margin-top: 30px;">
                    Met vriendelijke groet,<br><strong>G4S Operation Support Team</strong>
                  </p>
                </div>
              </div>
            `;
            
            try {
              MailApp.sendEmail({ to: sendToMail, subject: "Herhaling Toets: " + toetsNaam, htmlBody: emailHtml });
              hSheet.getRange(j+1, 4).setValue("Verstuurd");
              writeLog_("HERHALING_VERSTUURD", "Naar: " + sendToMail + " (Kandidaat: " + trueUpn + ") - Toets: " + toetsNaam);
            } catch(err) {
              hSheet.getRange(j+1, 4).setValue("Fout bij verzenden");
              Logger.log("Fout bij verzenden herinnering naar " + sendToMail + ": " + err.message);
            }
          }
        }
      }
    }
    
    props.setProperty('lastReminderCheck', todayStr);
    
    if (mutatieGedaan) {
      clearSystemCaches();
    }
    
  } catch(e) {
    Logger.log("Fout in checkEnVerstuurHerinneringen: " + e.message);
  }
}

function getGlobalCacheVersion() {
  var props = PropertiesService.getScriptProperties();
  var version = props.getProperty('GLOBAL_CACHE_VERSION');
  if (!version) {
    version = new Date().getTime().toString();
    props.setProperty('GLOBAL_CACHE_VERSION', version);
  }
  return version;
}

function clearSystemCaches() {
   try {
     PropertiesService.getScriptProperties().setProperty('GLOBAL_CACHE_VERSION', new Date().getTime().toString());
   } catch(e) {}
}

function putChunkedCache(key, value, expirationInSeconds) {
  var cache = CacheService.getScriptCache();
  var json = JSON.stringify(value);
  var chunkSize = 90000; 
  var chunks = [];
  for (var i = 0; i < json.length; i += chunkSize) {
    chunks.push(json.substring(i, i + chunkSize));
  }
  cache.put(key + '_count', chunks.length.toString(), expirationInSeconds);
  for (var j = 0; j < chunks.length; j++) {
    cache.put(key + '_' + j, chunks[j], expirationInSeconds);
  }
}

function getChunkedCache(key) {
  var cache = CacheService.getScriptCache();
  var count = cache.get(key + '_count');
  if (!count) return null;
  var json = '';
  for (var i = 0; i < parseInt(count); i++) {
    var chunk = cache.get(key + '_' + i);
    if (!chunk) return null; 
    json += chunk;
  }
  try {
    return JSON.parse(json);
  } catch(e) {
    return null; 
  }
}

function buildNightlyCache() {
   var email = "";
   try { email = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(e) {}
   var version = getGlobalCacheVersion();
   
   var dash = getDashboardData_calc_(email);
   putChunkedCache('dash_' + Utilities.base64Encode(email).substring(0,25) + '_' + version, dash, 21600); 
   
   var pb = getPowerBIData_calc_(email);
   putChunkedCache('pbi_' + Utilities.base64Encode(email).substring(0,25) + '_' + version, pb, 21600); 
}

function forceerHerinneringenHandmatig() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('lastReminderCheck'); 
  checkEnVerstuurHerinneringen(); 
  Logger.log("Geforceerde test voltooid!");
}

// ----------------------------------------------------------------------------------------------------
// 8. AANVRAGEN MODULE (FASE 2)
// ----------------------------------------------------------------------------------------------------

function submitNieuweAanvraag(data) {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase().trim();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    var shIdx = ss.getSheetByName(TAB_AANVRAAG_INDEX);
    if (!shIdx) {
      shIdx = ss.insertSheet(TAB_AANVRAAG_INDEX);
      shIdx.appendRow(["ID", "Naam", "Omschrijving", "Regio", "Klantnaam", "ManagerName", "ManagerEmail", "Norm", "Vragen", "Datum", "Aanvrager", "ToonAntwoorden"]);
    }
    
    var shDb = ss.getSheetByName(TAB_AANVRAAG_VRAGEN);
    if (!shDb) {
      shDb = ss.insertSheet(TAB_AANVRAAG_VRAGEN);
      shDb.appendRow(["ID", "Vraag", "Img", "O1", "O2", "O3", "O4", "Correct", "Verplicht", "Uitleg"]);
    }
    
    var newId = data.meta.id || ("AANV-" + Utilities.getUuid().slice(0, 8));
    
    shIdx.insertRowAfter(1);
    shIdx.getRange(2, 1, 1, 14).setValues([[
      newId, data.meta.naam, data.meta.desc || "", data.meta.regio || "", data.meta.klantnaam || "",
      data.meta.manName || "", data.meta.manEmail || "", data.meta.norm, data.vragen.length, new Date(), email, data.meta.toonAntwoorden !== false,
      data.meta.videoUrl || "", parseInt(data.meta.videoDuur, 10) || 0
    ]]);
    
    if (data.vragen.length > 0) {
      var newRows = data.vragen.map(function(q) {
        return [
          newId, q.vraag, q.img, q.o1, q.o2, q.o3, q.o4, q.correct, q.verplicht, q.uitleg || ""
        ];
      });
      var startRow = shDb.getLastRow() + 1;
      shDb.getRange(startRow, 1, newRows.length, 10).setValues(newRows);
    }

    if (data.meta.id) {
       deleteConceptAanvraag_(data.meta.id, email);
    }
    
    sendAanvraagNotificatieEmail_(data.meta.naam, email);
    writeLog_("TOETS_AANGEVRAAGD", "Aanvrager: " + email + ", Toets: " + data.meta.naam);
    return { success: true };
  } catch (e) {
    throw new Error("Aanvraag indienen mislukt: " + e.message);
  }
}

// ----------------------------------------------------------------------------------------------------
// 9. TUSSENTIJDS OPSLAAN LOGICA (CONCEPTEN)
// ----------------------------------------------------------------------------------------------------

function saveConceptAanvraag(data) {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase().trim();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    var shConcept = ss.getSheetByName(TAB_AANVRAAG_CONCEPT);
    if (!shConcept) {
      shConcept = ss.insertSheet(TAB_AANVRAAG_CONCEPT);
      shConcept.appendRow(["ID", "Aanvrager", "Datum", "Naam", "DataJSON"]);
    }
    
    var conceptId = data.meta.id || ("CONC-" + Utilities.getUuid().slice(0, 8));
    data.meta.id = conceptId; 
    var dataJson = JSON.stringify(data);
    
    var existingData = shConcept.getDataRange().getValues();
    var updated = false;
    
    for (var i = 1; i < existingData.length; i++) {
        if (existingData[i][0] === conceptId && String(existingData[i][1]).toLowerCase().trim() === email) {
            shConcept.getRange(i+1, 3, 1, 3).setValues([[new Date(), data.meta.naam || "Naamloos Concept", dataJson]]);
            updated = true;
            break;
        }
    }
    
    if (!updated) {
        shConcept.appendRow([conceptId, email, new Date(), data.meta.naam || "Naamloos Concept", dataJson]);
    }
    
    return { success: true, conceptId: conceptId, message: "Concept succesvol opgeslagen." };
  } catch (e) {
    throw new Error("Concept opslaan mislukt: " + e.message);
  }
}

function getConceptAanvragen() {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase().trim();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shConcept = ss.getSheetByName(TAB_AANVRAAG_CONCEPT);
    if (!shConcept) return [];
    
    var data = shConcept.getDataRange().getValues();
    var res = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase().trim() === email) {
          var dDate = data[i][2];
          var dStr = (dDate instanceof Date) ? dDate.toLocaleDateString('nl-NL') + " " + dDate.toLocaleTimeString('nl-NL', {hour:'2-digit', minute:'2-digit'}) : dDate;
          res.push({
             id: data[i][0],
             datum: dStr,
             naam: data[i][3] || "Naamloze Concept Toets",
             dataRaw: data[i][4]
          });
      }
    }
    return res.reverse();
  } catch (e) {
    throw new Error("Ophalen concepten mislukt: " + e.message);
  }
}

function loadConceptAanvraag(conceptId) {
    try {
        var email = Session.getActiveUser().getEmail().toLowerCase().trim();
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var shConcept = ss.getSheetByName(TAB_AANVRAAG_CONCEPT);
        if (!shConcept) throw new Error("Concepten tabblad ontbreekt.");
        
        var data = shConcept.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
            if (data[i][0] === conceptId && String(data[i][1]).toLowerCase().trim() === email) {
                return JSON.parse(data[i][4]);
            }
        }
        throw new Error("Concept niet gevonden of geen toegang.");
    } catch(e) {
        throw new Error("Laden concept mislukt: " + e.message);
    }
}

function deleteConceptAanvraag_(conceptId, userEmail) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shConcept = ss.getSheetByName(TAB_AANVRAAG_CONCEPT);
    if (!shConcept) return;
    
    var data = shConcept.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === conceptId && String(data[i][1]).toLowerCase().trim() === userEmail) {
            shConcept.deleteRow(i + 1);
            break;
        }
    }
}

// ----------------------------------------------------------------------------------------------------

function getAanvragenWachtrij() {
  try {
    var allowed = getAllowedToetsIds_();
    if (!allowed._isAdmin) throw new Error("Geen rechten voor de wachtrij.");
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shIdx = ss.getSheetByName(TAB_AANVRAAG_INDEX);
    if (!shIdx) return [];
    
    var data = shIdx.getDataRange().getValues();
    var res = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var dDate = data[i][9];
        var dStr = (dDate instanceof Date) ? dDate.toLocaleDateString('nl-NL') + " " + dDate.toLocaleTimeString('nl-NL', {hour:'2-digit', minute:'2-digit'}) : dDate;
        res.push({
          id: data[i][0], 
          naam: data[i][1], 
          regio: data[i][3], 
          klantnaam: data[i][4],
          vragen: data[i][8], 
          datum: dStr,
          aanvrager: data[i][10]
        });
      }
    }
    return res.reverse();
  } catch (e) {
    throw new Error("Wachtrij ophalen mislukt: " + e.message);
  }
}

function getAanvraagDetails(id) {
  try {
    var allowed = getAllowedToetsIds_();
    if (!allowed._isAdmin) throw new Error("Geen rechten.");
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shIdx = ss.getSheetByName(TAB_AANVRAAG_INDEX);
    var shDb = ss.getSheetByName(TAB_AANVRAAG_VRAGEN);
    
    if (!shIdx || !shDb) throw new Error("Aanvragen tabbladen niet gevonden.");
    
    var meta = null;
    var dataIdx = shIdx.getDataRange().getValues();
    for (var i = 1; i < dataIdx.length; i++) {
      if (dataIdx[i][0] === id) {
        meta = {
          id: dataIdx[i][0], naam: dataIdx[i][1], desc: dataIdx[i][2], regio: dataIdx[i][3], klantnaam: dataIdx[i][4],
          manName: dataIdx[i][5], manEmail: dataIdx[i][6], norm: dataIdx[i][7], aanvrager: dataIdx[i][10],
          toonAntwoorden: (dataIdx[i][11] !== false && String(dataIdx[i][11]).toUpperCase() !== "FALSE"),
          videoUrl: dataIdx[i][12] ? String(dataIdx[i][12]).trim() : "",
          videoDuur: parseInt(dataIdx[i][13], 10) || 0
        };
        break;
      }
    }
    if (!meta) throw new Error("Aanvraag niet gevonden.");
    
    var vragen = [];
    var dataDb = shDb.getDataRange().getValues();
    var dataDbDisplay = shDb.getDataRange().getDisplayValues();
    for (var j = 1; j < dataDb.length; j++) {
      if (dataDb[j][0] === id) {
        vragen.push({
          vraag: String(dataDbDisplay[j][1] || ""), img: String(dataDb[j][2] || ""), o1: String(dataDbDisplay[j][3] || ""), 
          o2: String(dataDbDisplay[j][4] || ""), o3: String(dataDbDisplay[j][5] || ""), o4: String(dataDbDisplay[j][6] || ""),
          correct: String(dataDb[j][7] || ""), verplicht: dataDb[j][8] === true, uitleg: String(dataDbDisplay[j][9] || "")
        });
      }
    }
    return { meta: meta, vragen: vragen };
  } catch (e) {
    throw new Error("Aanvraag details ophalen mislukt: " + e.message);
  }
}

function verwerkAanvraag(id, actie) {
  try {
    var allowed = getAllowedToetsIds_();
    if (!allowed._isAdmin) throw new Error("Geen rechten om aanvragen te verwerken.");
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shIdx = ss.getSheetByName(TAB_AANVRAAG_INDEX);
    var shDb = ss.getSheetByName(TAB_AANVRAAG_VRAGEN);
    
    if (!shIdx || !shDb) throw new Error("Aanvragen tabbladen niet gevonden.");
    
    var dataIdx = shIdx.getDataRange().getValues();
    var metaRow = null;
    var rowIndex = -1;
    
    for (var i = dataIdx.length - 1; i >= 1; i--) {
      if (dataIdx[i][0] === id) {
        metaRow = dataIdx[i];
        rowIndex = i + 1;
        break;
      }
    }
    
    if (!metaRow) throw new Error("Aanvraag niet gevonden.");
    
    var dataDb = shDb.getDataRange().getValues();
    var vragenRows = [];
    for (var j = dataDb.length - 1; j >= 1; j--) {
      if (dataDb[j][0] === id) {
        vragenRows.push(dataDb[j]);
        shDb.deleteRow(j + 1); 
      }
    }
    shIdx.deleteRow(rowIndex); 
    
    if (actie === 'AFKEUREN') {
      vragenRows.forEach(function(r) {
        if (r[2] && String(r[2]).indexOf("drive.google.com") > -1) {
          deleteImageFromDrive(r[2]);
        }
      });
      writeLog_("AANVRAAG_AFGEKEURD", "Aanvraag: " + metaRow[1] + " (Aanvrager: " + metaRow[10] + ")");
      return { success: true, message: "Aanvraag is afgekeurd en definitief verwijderd." };
    } 
    else if (actie === 'GOEDKEUREN') {
      var liveIdx = ss.getSheetByName(TAB_INDEX);
      var liveDb = ss.getSheetByName(TAB_DB);
      
      var newLiveId = Utilities.getUuid().slice(0, 8);
      
      liveIdx.insertRowBefore(2);
      liveIdx.getRange(2, 1, 1, 14).setValues([[
        newLiveId, metaRow[1], metaRow[2], metaRow[3], metaRow[4], metaRow[5], metaRow[6], metaRow[7], metaRow[8], new Date(), true, metaRow[11] !== false,
        metaRow[12] || "", parseInt(metaRow[13], 10) || 0
      ]]);
      
      if (vragenRows.length > 0) {
        var liveVragen = vragenRows.reverse().map(function(r) {
          return [newLiveId, r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9]];
        });
        var startLiveRow = liveDb.getLastRow() + 1;
        liveDb.getRange(startLiveRow, 1, liveVragen.length, 10).setValues(liveVragen);
      }
      
      writeLog_("AANVRAAG_GOEDGEKEURD", "Aanvraag gepromoveerd naar live toets: " + metaRow[1]);
      
      sendAanvraagGoedgekeurdEmail_(metaRow[1], metaRow[10]);

      var rawManagers = String(metaRow[6]).toLowerCase().trim();
      var eersteManagerEmail = rawManagers.split(/[\n,;]+/)[0].trim();

      if (eersteManagerEmail && eersteManagerEmail.indexOf('@') > -1) {
        var shRechten = ss.getSheetByName(TAB_RECHTEN);
        if (!shRechten) {
          shRechten = ss.insertSheet(TAB_RECHTEN);
          shRechten.appendRow(["Email", "Rol"]);
        }
        var rechtenData = shRechten.getDataRange().getValues();
        var hasRole = false;
        for (var r = 1; r < rechtenData.length; r++) {
          if (String(rechtenData[r][0]).toLowerCase().trim() === eersteManagerEmail) {
            hasRole = true;
            break;
          }
        }
        if (!hasRole) {
          shRechten.appendRow([eersteManagerEmail, "Toetsenmaker"]);
          writeLog_("GEBRUIKER_AANGEMAAKT", "Email: " + eersteManagerEmail + ", Rol: Toetsenmaker (Automatisch via aanvraag)");
        }
      }
      
      clearSystemCaches();
      return { success: true, message: "Aanvraag is succesvol goedgekeurd en verplaatst naar de Toetsen Index!" };
    }
  } catch (e) {
    throw new Error("Aanvraag verwerken mislukt: " + e.message);
  }
}

function importeerAanvraagVanuitTemplate(url) {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase().trim();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var src = SpreadsheetApp.openByUrl(url).getSheets()[0];
    
    if (src.getLastRow() < 8) {
       throw new Error("Validatie mislukt: Er zijn geen vragen gevonden in deze sheet (vanaf rij 8). Gebruik het officiële G4S Aanleverformulier.");
    }
    
    var meta = src.getRange(4, 1, 2, 11).getValues();
    var toetsNaam = meta[0][2];
    
    if (!toetsNaam || String(toetsNaam).trim() === "") {
       throw new Error("Validatie mislukt: De toetsnaam ontbreekt op de verwachte locatie (Cel C4). Controleer of dit het officiële G4S Aanleverformulier is.");
    }
    
    var newId = "AANV-" + Utilities.getUuid().slice(0, 8);
    var norm = parseNorm(meta[1][6]);
    
    var shIdx = ss.getSheetByName(TAB_AANVRAAG_INDEX);
    if (!shIdx) {
      shIdx = ss.insertSheet(TAB_AANVRAAG_INDEX);
      shIdx.appendRow(["ID", "Naam", "Omschrijving", "Regio", "Klantnaam", "ManagerName", "ManagerEmail", "Norm", "Vragen", "Datum", "Aanvrager", "ToonAntwoorden"]);
    }
    
    var idxRow = [
      newId, 
      toetsNaam,  
      meta[1][2] || "", 
      meta[0][4] || "", 
      meta[1][4] || "", 
      meta[0][8] || "", 
      meta[1][8] || "", 
      norm,        
      meta[0][6] || 0, 
      new Date(),
      email,
      true 
    ];
    
    shIdx.insertRowAfter(1);
    shIdx.getRange(2, 1, 1, 12).setValues([idxRow]);
    
    var vr = src.getRange(8, 1, src.getLastRow() - 7, 12).getValues();
    var vrDisplay = src.getRange(8, 1, src.getLastRow() - 7, 12).getDisplayValues();
    var dbRows = [];
    
    for (var i = 0; i < vr.length; i++) {
      if (!vr[i][2]) { break; } 
      dbRows.push([
        newId, 
        String(vrDisplay[i][2]).trim(),  
        String(vrDisplay[i][3]).trim(),  
        vrDisplay[i][4],                 
        vrDisplay[i][5],                 
        vrDisplay[i][6],                 
        vrDisplay[i][7],                 
        normalizeCorrectIdx_(vr[i][8]),                     
        (String(vr[i][9]).toLowerCase().indexOf("ja") > -1), 
        vrDisplay[i][10]                 
      ]);
    }
    
    if (dbRows.length > 0) {
      var shDb = ss.getSheetByName(TAB_AANVRAAG_VRAGEN);
      if (!shDb) {
        shDb = ss.insertSheet(TAB_AANVRAAG_VRAGEN);
        shDb.appendRow(["ID", "Vraag", "Img", "O1", "O2", "O3", "O4", "Correct", "Verplicht", "Uitleg"]);
      }
      var startRow = shDb.getLastRow() + 1;
      shDb.getRange(startRow, 1, dbRows.length, 10).setValues(dbRows);
    }
    
    sendAanvraagNotificatieEmail_(toetsNaam, email);
    writeLog_("TOETS_AANGEVRAAGD_VIA_IMPORT", "Aanvrager: " + email + ", Toets: " + toetsNaam);
    return { success: true, message: "Aanvraag succesvol geïmporteerd en ingediend (" + dbRows.length + " vragen)." };
  } catch(e) { 
    throw new Error("Importeren van aanvraag mislukt: " + e.message); 
  }
}

function sendAanvraagNotificatieEmail_(toetsNaam, aanvrager) {
  try {
    var scriptUrl = ScriptApp.getService().getUrl();
    var savedUrl = PropertiesService.getScriptProperties().getProperty('SAVED_WEBAPP_URL');
    if (savedUrl && scriptUrl.indexOf("/dev") > -1) { scriptUrl = savedUrl; }
    scriptUrl = scriptUrl.replace(/\/dev$/i, "/exec");

    var logoUrl = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : "https://www.g4s.com/-/media/g4s/global/images/logos/g4s---an-allied-universal-company-logo/g4s_one-company_logo.png";
    var emailHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f6f9; padding: 20px; border-radius: 8px;">
        <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; border-top: 4px solid #D32F2F; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px;">
            <tr>
              <td style="vertical-align: middle;"><h2 style="color: #232323; font-size: 20px; margin: 0;">Nieuwe Toets Aanvraag</h2></td>
              <td style="text-align: right; vertical-align: middle;"><img src="${logoUrl}" alt="G4S Logo" style="height: 50px;"></td>
            </tr>
          </table>
          <p style="color: #333; font-size: 15px; line-height: 1.6;">Beste Operation Support Team,</p>
          <p style="color: #333; font-size: 15px; line-height: 1.6;">Er is zojuist een nieuwe kennistoets ingediend ter goedkeuring via het Self-Service Portaal.</p>
          
          <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0; border: 1px solid #eee;">
            <p style="margin: 0 0 8px 0; color: #333;"><strong>Aanvrager:</strong> ${aanvrager}</p>
            <p style="margin: 0; color: #333;"><strong>Naam Toets:</strong> ${toetsNaam}</p>
          </div>
          <div style="text-align: center; margin: 30px 0;">
            <a href="${scriptUrl}" target="_blank" rel="noopener noreferrer" style="background-color: #D32F2F; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; text-transform: uppercase;">OPEN PORTAAL & BEOORDEEL</a>
          </div>
          <p style="color: #666; font-size: 13px; line-height: 1.5; margin-top: 30px;">
            Dit is een automatisch gegenereerd bericht vanuit het G4S Kennistoetsen systeem.
          </p>
        </div>
      </div>
    `;
    try {
      MailApp.sendEmail({ 
          to: (typeof CONFIG !== 'undefined') ? CONFIG.SUPPORT_EMAIL : "operation.supportteam@nl.g4s.com", 
          subject: "Actie vereist: Nieuwe Toets Aanvraag - " + toetsNaam, 
          htmlBody: emailHtml 
      });
    } catch(err) {
      Logger.log("Fout bij verzenden notificatie naar OST: " + err.message);
    }
  } catch(e) {
    Logger.log("Fout in sendAanvraagNotificatieEmail_: " + e.message);
  }
}

function sendAanvraagGoedgekeurdEmail_(toetsNaam, aanvrager) {
  try {
    var scriptUrl = ScriptApp.getService().getUrl();
    var savedUrl = PropertiesService.getScriptProperties().getProperty('SAVED_WEBAPP_URL');
    if (savedUrl && scriptUrl.indexOf("/dev") > -1) { scriptUrl = savedUrl; }
    scriptUrl = scriptUrl.replace(/\/dev$/i, "/exec");

    var logoUrl = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : "https://www.g4s.com/-/media/g4s/global/images/logos/g4s---an-allied-universal-company-logo/g4s_one-company_logo.png";
    var emailHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f6f9; padding: 20px; border-radius: 8px;">
        <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; border-top: 4px solid #2E7D32; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px;">
            <tr>
              <td style="vertical-align: middle;"><h2 style="color: #232323; font-size: 20px; margin: 0;">Toets Aanvraag Goedgekeurd</h2></td>
              <td style="text-align: right; vertical-align: middle;"><img src="${logoUrl}" alt="G4S Logo" style="height: 50px;"></td>
            </tr>
          </table>
          <p style="color: #333; font-size: 15px; line-height: 1.6;">Beste collega,</p>
          <p style="color: #333; font-size: 15px; line-height: 1.6;">Goed nieuws! Je aangevraagde kennistoets <strong>${toetsNaam}</strong> is zojuist beoordeeld en <strong>goedgekeurd</strong> door het Operation Support Team.</p>
          
          <div style="background-color: #e8f5e9; border-left: 4px solid #2e7d32; padding: 15px; margin: 20px 0; border: 1px solid #c8e6c9;">
            <p style="margin: 0; color: #1b5e20; font-size: 14px; line-height: 1.5;"><strong>Wat moet je nu nog doen?</strong><br><br>
            Je hebt vanaf nu toegang gekregen tot de beheeromgeving van deze toets. Het is belangrijk om eerst de <strong>handleiding</strong> in de applicatie goed door te nemen.<br><br>
            Om de toets daadwerkelijk uit te zetten naar je kandidaten, moet je zelf de e-mailadressen (G4S UPN) nog invoeren in de module <em>'Uitnodigingen Versturen'</em>.</p>
          </div>

          <div style="text-align: center; margin: 30px 0;">
            <a href="${scriptUrl}" target="_blank" rel="noopener noreferrer" style="background-color: #D32F2F; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; text-transform: uppercase;">OPEN PORTAAL & VERSTUUR</a>
          </div>
          <p style="color: #666; font-size: 13px; line-height: 1.5; margin-top: 30px;">
            Succes met het uitzetten van de toets!<br><strong>G4S Operation Support Team</strong>
          </p>
        </div>
      </div>
    `;
    try {
      MailApp.sendEmail({ 
          to: aanvrager, 
          subject: "Toets Goedgekeurd: " + toetsNaam, 
          htmlBody: emailHtml 
      });
    } catch (err) {
      Logger.log("Fout bij verzenden goedkeuringsmail naar " + aanvrager + ": " + err.message);
    }
  } catch(e) {
    Logger.log("Fout in sendAanvraagGoedgekeurdEmail_: " + e.message);
  }
}

function exportNogTeDoenSheet(filterToetsNaam, clientEmail) {
  try {
    var allowedIds = getAllowedToetsIds_(clientEmail);
    var email = getViewAsEmail_(clientEmail);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shSes = ss.getSheetByName(TAB_SESSIES);
    var shIdx = ss.getSheetByName(TAB_INDEX);
    
    if (!shSes || !shIdx) {
      throw new Error("Nodige tabbladen (Sessie_Log of Toets_Index) ontbreken.");
    }
    
    var idxData = shIdx.getDataRange().getValues();
    var toetsMap = {};
    for (var i = 1; i < idxData.length; i++) {
      toetsMap[idxData[i][0]] = {
        naam: idxData[i][1],
        manager: idxData[i][5] || "-"
      };
    }

    var sesData = shSes.getDataRange().getValues();
    var exportData = [["Datum Uitgenodigd", "Toets Naam", "Verstuurd Naar (E-mailadres)", "Manager", "Status"]];
    var filterList = Array.isArray(filterToetsNaam)
      ? filterToetsNaam.map(function(s){ return String(s).trim().toLowerCase(); }).filter(Boolean)
      : (filterToetsNaam ? [String(filterToetsNaam).trim().toLowerCase()] : []);
    
    for (var j = 1; j < sesData.length; j++) {
      if (sesData[j][10] === "OPEN") {
        var tId = sesData[j][1];
        
        if (!allowedIds._isAdmin && !allowedIds[tId]) {
          continue;
        }
        
        var tNaam = toetsMap[tId] ? toetsMap[tId].naam : "Onbekende Toets";
        
        if (filterList.length > 0 && filterList.indexOf(String(tNaam).trim().toLowerCase()) === -1) {
          continue;
        }
        
        var rawDatum = sesData[j][5];
        var dStr = "-";
        if (rawDatum instanceof Date) {
           dStr = rawDatum.toLocaleDateString('nl-NL') + " " + rawDatum.toLocaleTimeString('nl-NL', {hour: '2-digit', minute:'2-digit'});
        } else if (rawDatum) {
           dStr = String(rawDatum);
        }
        
        var managerNaam = toetsMap[tId] ? toetsMap[tId].manager : "-";
        
        // Verstuurd Naar kolom (Kolom N is index 13) of fallback naar Deelnemer UPN (Kolom I is index 8)
        var verstuurdNaar = sesData[j][13] || sesData[j][8] || "-";
        
        exportData.push([
          dStr,
          tNaam,
          verstuurdNaar,
          managerNaam,
          "Openstaand"
        ]);
      }
    }
    
    if (exportData.length === 1) {
       throw new Error("Er zijn geen openstaande uitnodigingen gevonden voor de gekozen filters.");
    }
    
    var dNu = new Date();
    var dateString = Utilities.formatDate(dNu, "Europe/Amsterdam", "dd_MM_yyyy_HHmm");
    var safeToetsPrefix = filterToetsNaam ? String(filterToetsNaam).replace(/[\\/:*?"<>| ]/g, "_") + "_" : "";
    var sheetName = "Export_NogTeDoen_" + safeToetsPrefix + dateString;
    
    var newSpreadsheet = SpreadsheetApp.create(sheetName);
    var sheet = newSpreadsheet.getActiveSheet();
    sheet.setName("Openstaand");
    
    sheet.getRange(1, 1, exportData.length, exportData[0].length).setValues(exportData);
    sheet.getRange(1, 1, 1, exportData[0].length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    
    try {
      sheet.autoResizeColumns(1, exportData[0].length);
    } catch(eResize) {
      Logger.log("autoResizeColumns waarschuwing: " + eResize.message);
    }
    
    var file = DriveApp.getFileById(newSpreadsheet.getId());
    
    // Stel algemene toegang strikt in op 'Beperkt' (PRIVATE) zodat niet iedereen binnen het domein toegang heeft
    try {
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    } catch(ePriv) {
      Logger.log("setSharing PRIVATE waarschuwing: " + ePriv.message);
    }
    
    // Alleen als de export wordt opgevraagd door een andere gebruiker dan de eigenaar,
    // voegen we uitsluitend die specifieke gebruiker toe als lezer
    var effectiveUser = "";
    try {
      effectiveUser = Session.getEffectiveUser().getEmail().toLowerCase().trim();
    } catch(eEff) {}
    var currentUser = getCurrentUserEmail_();
    var requestingUser = (clientEmail || currentUser || "").toLowerCase().trim();
    
    if (requestingUser && effectiveUser && requestingUser !== effectiveUser && requestingUser.indexOf("@") > -1) {
      try {
        file.addViewer(requestingUser);
      } catch(eView) {
        Logger.log("addViewer waarschuwing: " + eView.message);
      }
    }
    
    writeLog_("EXPORT_NOGTEDOEN", "Gebruiker heeft openstaande acties geëxporteerd. Filter: " + (filterToetsNaam || "Alle"));
    
    return { success: true, url: newSpreadsheet.getUrl() };
  } catch(e) {
    var errMsg = e.message || String(e);
    if (errMsg.indexOf("Er zijn geen openstaande") > -1) {
      throw new Error(errMsg);
    }
    throw new Error("Fout bij aanmaken export: " + errMsg);
  }
}