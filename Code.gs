function doGet(e) {
  let temp = 'Index';
  if ('page' in e.parameters) {
    temp = e.parameters['page'][0];
  }

  // Check if cached data is available
  let cachedData = CacheService.getScriptCache().get('cachedData');
  if (!cachedData) {
    // If not, fetch data from sheets
    var gniazdo = getGniazdo();
    var stan = getStatus();
    var pracownik = getPracownik();
    var pracownik1 = getPracownik();
    var zlecenie = getZlecenie();

    // Cache the data for 5 minutes (adjust as needed)
    CacheService.getScriptCache().put('cachedData', JSON.stringify({ gniazdo, stan, pracownik, pracownik1, zlecenie }), 15000); // Reduced caching time
  } else {
    // If cached data is available, use it
    cachedData = JSON.parse(cachedData);
    var gniazdo = cachedData.gniazdo;
    var stan = cachedData.stan;
    var pracownik = cachedData.pracownik;
    var pracownik1 = cachedData.pracownik1;
    var zlecenie = cachedData.zlecenie;
  }

  try {
    const html = HtmlService.createTemplateFromFile(temp);
    html.message = '';
    html.gniazdo = gniazdo;
    html.stan = stan;
    html.pracownik = pracownik;
    html.data = { title: temp, e: e };
    html.pracownik1 = pracownik1;
    html.zlecenie = zlecenie;
    return html.evaluate();
  } catch (err) {
    const html = HtmlService.createHtmlOutput('Page not found 404 Error: ' + JSON.stringify(err));
    return html;
  }
}

function getChartData(){
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gantt");
  const data = ss.getDataRange().getValues();
  return data;
}


function getScriptUrl(){
  const url = ScriptApp.getService().getUrl();
  return url;
}

/////////////////////////// EMAIL ///////////////////////////////////////////

function sendMyEmail(recipient, section, additionalInfo) {
  let email = '';
  let subject = '';

  switch (section) {

    case 'GIĘTARKI':
      subject = 'Wezwanie na gniazdo GIĘTAREK';
      break;
    case 'AUTOMAT ARTSHAPE':
      subject = 'Wezwanie na gniazdo AUTOMAT ARTSHAPE';
      break;
    case 'AUTOMAT GNĄCY':
      subject = 'Wezwanie na gniazdo AUTOMAT GNĄCY';
      break;
    case 'GIĘTARKI':
      subject = 'Wezwanie na gniazdo GIĘTARKI';
      break;
    case 'NOŻYCA GILOTYNOWA':
      subject = 'Wezwanie na gniazdo NOŻYCA GILOTYNOWA';
      break;
    case 'PLOTER':
      subject = 'Wezwanie na gniazdo PLOTER';
      break;
    case 'SPAWARKI CO2':
      subject = 'Wezwanie na gniazdo SPAWARKI CO2';
      break;
    case 'SZLIFIERKA':
      subject = 'Wezwanie na gniazdo SZLIFIERKA';
      break;
    case 'TARCZOWA':
      subject = 'Wezwanie na gniazdo TARCZOWA';
      break;
    case 'WIERTARKA STOŁOWA':
      subject = 'Wezwanie na gniazdo WIERTARKA STOŁOWA';
      break;
    case 'WYKRAWARKI':
      subject = 'Wezwanie na gniazdo WYKRAWARKI';
      break;
    case 'ZGRZEWARKI':
      subject = 'Wezwanie na gniazdo ZGRZEWARKI';
      break;



    default:
      // Default case
      subject = 'Default Subject';
  }

  switch (recipient) {
    case 'Brygadzista':
      email = 'dorianduda92@gmail.com';
      break;
    case 'UR':
      email = 'dorian.duda@luxiona.com';
      break;
    case 'Konstrukcyjny':
      email = 'dorian.duda@luxiona.com';
      break;
    case 'Jakość':
      email = 'dorian.duda@luxiona.com';
      break;
    case 'KP':
      email = 'jaroslaw.sadzinski@luxiona.com';
      break;

    default:
      // Default case
      email = '';
  }


  var body = 'wiadomosc wygenerowana automatycznie';
    if (additionalInfo) {
    body += '\nDodatkowe informacje: ' + additionalInfo;
  }
  MailApp.sendEmail(email, subject, body);
}




///////////////////////////////////////////////////////////////////////////////////////


function howMuchEmailsLeft(){
var quotaLeft = MailApp.getRemainingDailyQuota();
Logger.log(quotaLeft);
}

    function dataSevedSuccessfully() {
        Swal.fire({
            icon: 'success',
            title: 'Info',
            text: 'Dane zapisane poprawnie!'
        })
    }


function menu(){
  const url = getScriptUrl();
  let html = HtmlService.createHtmlOutputFromFile ('menu').getContent();
  html = html.replace(/\?page/g, url+'?page');
  //Logger.log(html);
  return html;
}

function sidebar_vertical(){
  const url = getScriptUrl();
  let html = HtmlService.createHtmlOutputFromFile ('sidebar_vertical').getContent();
  html = html.replace(/\?page/g, url+'?page');
  //Logger.log(html);
  return html;
}


//GET DATA FROM GOOGLE SHEET AND RETURN AS AN ARRAY
function getData() {
  var spreadSheetId = "1q4k9diFgLQbJoTrcRvTIfcjLispMvhJi_iSjt9H-Wb8"; //CHANGE
  var dataRange = "Data!A2:H"; //CHANGE

  var range = Sheets.Spreadsheets.Values.get(spreadSheetId, dataRange);
  var values = range.values;

  return values;
}


//GET DATA FROM GOOGLE SHEET AND RETURN AS AN ARRAY
function getDataKomunikaty() {
  var spreadSheetId = "1q4k9diFgLQbJoTrcRvTIfcjLispMvhJi_iSjt9H-Wb8"; //CHANGE
  var dataRange = "Komunikaty2!A2:D"; //CHANGE

  var range = Sheets.Spreadsheets.Values.get(spreadSheetId, dataRange);
  var values = range.values;

  return values;
}


function doPost(e) {
  
//  Logger.log(JSON.stringify(e));
  
  var gniazdo = e.parameters.gniazdo.toString();
  var zlecenie = e.parameters.zlecenie.toString();
  var stan = e.parameters.stan.toString();
  var pracownik = e.parameters.pracownik.toString();
  var pracownik1 = e.parameters.pracownik.toString();


  AddRecord(gniazdo, zlecenie, stan, pracownik, pracownik1);
  
  var htmlOutput =  HtmlService.createTemplateFromFile('Index');
  var gniazdo = getGniazdo();
  var stan = getStatus();
  var pracownik = getPracownik();
  var pracownik1 = getPracownik();
  var zlecenie = getZlecenie();

  htmlOutput.message = 'Record Added';
  htmlOutput.gniazdo = gniazdo;
  htmlOutput.stan = stan;
  htmlOutput.pracownik = pracownik;
  htmlOutput.pracownik1 = pracownik1;
  htmlOutput.zlecenie = zlecenie;

  return htmlOutput.evaluate();
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function processForm(formObject){
  var url="https://docs.google.com/spreadsheets/d/1q4k9diFgLQbJoTrcRvTIfcjLispMvhJi_iSjt9H-Wb8/edit?pli=1#gid=0";
  var ss= SpreadsheetApp.openByUrl(url);
  var ws=ss.getSheetByName("Data");

  ws.appendRow([
    formObject.obszar,
    formObject.numer_pracownika1,
    formObject.numer_zlecenia,
    formObject.status,
    formObject.liczba_sztuk,
    formObject.komentarz,
    new Date(),
    formObject.numer_pracownika2,
  ]);
}

function getSheetData()  {
  var a= SpreadsheetApp.getActiveSpreadsheet();
  var b = a.getSheetByName('Data');
  var c = b.getDataRange();
  return c.getValues();
}

function getSheetData1() {
  var x = SpreadsheetApp.getActiveSpreadsheet();
  var y = x.getSheetByName('Komunikaty');
  var z = y.getRange('A1:F10');
  return z.getValues();
}
/////////////////////////////////////////////////////////////////////////////////////////////////////

// function getGniazdo() {
//  var ss = SpreadsheetApp.openById("1GfARKhFtA1F5s3V6-NuRUA96nN78UKpy16-tmHvuLfw");
//  var lovSheet = ss.getSheetByName("ZLECENIA");
//  var dataRange = lovSheet.getRange(2, 1, lovSheet.getLastRow() - 1, 1);
//  var data = dataRange.getValues();
//  var return_array = Array.from(new Set(data.flat())); // Using Set to get unique values
//  return return_array;
//}


function getGniazdo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lovSheet = ss.getSheetByName("OBSADA");
  var getLastRow = lovSheet.getLastRow();
  var uniqueValues = new Set();

  for (var i = 2; i <= getLastRow; i++) {
    uniqueValues.add(lovSheet.getRange(i, 1).getValue());
  }

  return Array.from(uniqueValues);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
function getZlecenie(gniazdo) {
  var ss = SpreadsheetApp.openById("1dPbJaPShM2GJ2D5JQshsYCZWXj6UZBpELyjAp8Rf-6s");
  var lovSheet = ss.getSheetByName("ZLECENIA");
  var dataRange = lovSheet.getRange(2, 1, lovSheet.getLastRow() - 1, 4); // Adjust the range accordingly
  var data = dataRange.getValues();

  var return_array = data
    .filter(function (row) {
      return row[0] === gniazdo;
    })
    .map(function (row) {
      return { numerZlecenia: row[1], nazwaCzesci: row[2], nazwaOperacji: row[3] };
    });

  return return_array;
}

////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////////
function getStatus() { 
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var lovSheet = ss.getSheetByName("STATUS"); 
  var getLastRow = lovSheet.getLastRow();
  var return_array = [];
  for(var i = 2; i <= getLastRow; i++)
  {
      if(return_array.indexOf(lovSheet.getRange(i, 1).getValue()) === -1) {
        return_array.push(lovSheet.getRange(i, 1).getValue());
      }
  }
  return return_array;  
}

////////////////////////////////////////////////////////////////////////////////////////////////////
function getPracownik() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lovSheet = ss.getSheetByName("OBSADA");
  var getLastRow = lovSheet.getLastRow();
  var uniqueValues = new Set();

  for (var i = 2; i <= getLastRow; i++) {
    uniqueValues.add(lovSheet.getRange(i, 2).getValue());
  }

  return Array.from(uniqueValues);
}

///////////////////////////////////////////////////////////////////////////////////////////////////

function AddRecord(gniazdo, zlecenie, stan, pracownik, pracownik1) {
  var url = '';   //URL OF GOOGLE SHEET;
  var ss= SpreadsheetApp.openByUrl(url);
  var dataSheet = ss.getSheetByName("Data");
  dataSheet.appendRow([gniazdo, zlecenie, stan, pracownik, new Date(), pracownik1]);
}

function getUrl() {
 var url = ScriptApp.getService().getUrl();
 return url;
}
