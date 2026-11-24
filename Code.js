// ====================================================================================================
// G4S KENNISTOETSEN - CONFIGURATIE & ROUTING (WEB APP VERSIE)
// ====================================================================================================

// 1. CENTRALE CONFIGURATIE
// Pas deze ID's en instellingen aan wanneer de applicatie naar een andere omgeving wordt verplaatst.
var CONFIG = {
  SPREADSHEET_ID: "1egy6OZVcMm_4hDyVFlrAbaOgRAqjiC6-mZCsioCuZcM",
  DRIVE_FOLDER_ID: "1CGktm8Hro1kbZVARuMAAkhoJUUk_xQKa",
  VIDEO_FOLDER_ID: "1_Jq--xDRoTTJlr_b5kH2j3LOT6tNU6Mc",
  LOGO_ID: "1OFSDHFxHGqg8EKcLu8qtRtqSRiXwl89c",
  LOGO_URL: "https://www.g4s.com/-/media/g4s/global/images/logos/g4s---an-allied-universal-company-logo/g4s_one-company_logo.png",
  SUPPORT_EMAIL: "operation.supportteam@nl.g4s.com",
  SUPPORT_EMAIL_ALT: "operation.supporteam@nl.g4s.com"
};

// 2. GLOBALE VARIABELEN (Gekoppeld aan configuratie)
var SPREADSHEET_ID = CONFIG.SPREADSHEET_ID;

// Tabblad namen
var TAB_INDEX = "Toets_Index";
var TAB_DB = "Vragen_Database";
var TAB_RESULTATEN = "Resultaten";
var TAB_SESSIES = "Sessie_Log";
var TAB_HERINNERINGEN = "Herinneringen";
var TAB_REGIO = "Regio Indeling";
var TAB_KLANTEN = "Klanten Lijst";
var TAB_MUTATIE = "Mutatie_Log";
var TAB_EMP = "Dropdown Info";
var TAB_AANVULLING = "Account aanvulling";
var TAB_RECHTEN = "Beheer_Rechten"; 
var TAB_AANVRAAG_INDEX = "Aanvragen_Index"; 
var TAB_AANVRAAG_VRAGEN = "Aanvragen_Vragen"; 
var TAB_AANVRAAG_CONCEPT = "Aanvragen_Concept";
var TAB_POGINGEN = "Quiz_Pogingen";

function doGet(e) {
  var tId = "";
  var sId = "";
  if (e && e.parameter) {
     tId = e.parameter.toetsId || "";
     sId = e.parameter.sid || "";
  }

  var mode = (tId !== "") ? "TOETS" : "BEHEER";
  var userEmail = "";
  try {
     userEmail = Session.getActiveUser().getEmail() || "";
  } catch(err) {}

  var isAdmin = false;
  try {
    if (mode === "BEHEER") {
       isAdmin = getAllowedToetsIds_(userEmail)._isAdmin;
    }
  } catch(err) {}

  var isMaintenance = false;
  try {
    isMaintenance = PropertiesService.getScriptProperties().getProperty('MAINTENANCE_MODE') === 'true';
  } catch(err) {}

  var tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.mode = mode;
  tmpl.toetsId = tId;
  tmpl.sid = sId;
  tmpl.userEmail = userEmail;
  tmpl.isAdmin = (isAdmin === true) ? true : false;
  tmpl.isMaintenance = isMaintenance;
  tmpl.logoUrl = CONFIG.LOGO_URL;

  return tmpl.evaluate()
     .setTitle('Kennistoetsen G4S')
     .setFaviconUrl('https://www.g4s.com/favicon.ico')
     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
     .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====================================================================================================
// HELPER FUNCTIES & ONDERHOUDSMODUS
// ====================================================================================================

function parseNorm(val) {
  if (!val) { return 70; }
  var str = String(val).replace('%', '').replace(',', '.');
  var num = parseFloat(str);
  if (isNaN(num)) { return 70; }
  if (num > 0 && num <= 1) { num = num * 100; }
  return Math.round(num);
}

function parseScoreStr_(str) {
  var correct = 0;
  var total = 0;
  var percent = 0;
  if (!str) { return { correct: 0, total: 0, percent: 0 }; }
  var parts = String(str).split(' ');
  if (parts.length > 0) {
    var frac = parts[0].split('/');
    if (frac.length === 2) {
      correct = parseInt(frac[0], 10) || 0;
      total = parseInt(frac[1], 10) || 0;
    }
  }
  if (parts.length > 1) {
    percent = parseFloat(parts[1].replace('%', '').replace(',', '.')) || 0;
  }
  return { correct: correct, total: total, percent: percent };
}

function normalizeCorrectIdx_(val) {
  if (val === null || val === undefined) { return ""; }
  return String(val).trim().replace(/\./g, ',');
}

function getMaintenanceMode() {
  try {
    return PropertiesService.getScriptProperties().getProperty('MAINTENANCE_MODE') === 'true';
  } catch(e) { 
    return false; 
  }
}

function toggleMaintenanceMode(state) {
  var allowed = getAllowedToetsIds_();
  if (!allowed._isAdmin) {
    throw new Error("Geen rechten voor deze actie.");
  }
  PropertiesService.getScriptProperties().setProperty('MAINTENANCE_MODE', state ? 'true' : 'false');
  return state;
}