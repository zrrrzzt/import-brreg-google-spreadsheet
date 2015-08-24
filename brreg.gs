/**
 *
 * Import data from brreg
 * @customfunction
 **/

var pageNum = -1;
var batchSize = 100;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BRREG')
    .addItem('Import from brreg', 'importSetup')
    .addToUi();
}

function gotNext(links){
  var next = false;
  links.forEach(function(item){
    if(item.rel == 'next') {
      next = true;
    }
  });
  return next;
}

function addHeaders() {
  var out = [];
  var itemFields = getFields();
  itemFields.forEach(function(prop){
    out.push(prop.name);
  });
  writeData([out]);
}

function importSetup() {
  pageNum = -1;
  var sheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var inputKode = ui.prompt(
    'Add næringskode',
    ui.ButtonSet.OK_CANCEL);
  var inputStartdato = ui.prompt(
    'Add startdato',
    'Format: yyyy-MM-DD',
    ui.ButtonSet.OK_CANCEL);

  var kode = inputKode.getResponseText();
  var startdato = inputStartdato.getResponseText();
  var lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    addHeaders();
  }

  importBRREG(kode, startdato);
}

function importBRREG(kode, startdato) {
  pageNum++;
  var url = "http://data.brreg.no/enhetsregisteret/enhet.json?page=" + pageNum + "&size=" + batchSize + "&$filter=startswith(naeringskode1/kode, '" + kode + "')+and+registreringsdatoEnhetsregisteret+gt+datetime'" + startdato + "T00:00'";
  var result = UrlFetchApp.fetch(url);
  var json = JSON.parse(result.getContentText());
  var data = flattenData(json.data);

  writeData(data);

  if (gotNext(json.links)) {
    importBRREG(kode, startdato);
  }
}

function flattenData(list) {
  var data = [];
  list.forEach(function(item) {
    var row = unwrapObject(item);
    data.push(row);
  });
  return data;
}

function getFields() {
  var fields = [
    {name: 'Organisasjonsnummer', id: 'organisasjonsnummer'},
    {name: 'Navn', id: 'navn'},
    {name: 'Hjemmeside', id: 'hjemmeside'},
    {name: 'Næringskode', id: 'naeringskode1', sub:'kode'},
    {name: 'Beskrivelse', id: 'naeringskode1', sub:'beskrivelse'},
    {name: 'Postadresse', id: 'postadresse', sub:'adresse'},
    {name: 'Postnummer', id: 'postadresse', sub:'postnummer'},
    {name: 'Poststed', id: 'postadresse', sub:'poststed'},
    {name: 'Kommunenummer', id: 'postadresse', sub:'kommunenummer'},
    {name: 'Kommune', id: 'postadresse', sub:'kommune'},
    {name: 'Landkode', id: 'postadresse', sub:'landkode'},
    {name: 'Land', id: 'postadresse', sub:'land'},
    {name: 'Forretningsadresse', id: 'forretningsadresse', sub:'adresse'},
    {name: 'Postnummer', id: 'forretningsadresse', sub:'postnummer'},
    {name: 'Poststed', id: 'forretningsadresse', sub:'poststed'},
    {name: 'Kommunenummer', id: 'forretningsadresse', sub:'kommunenummer'},
    {name: 'Kommune', id: 'forretningsadresse', sub:'kommune'},
    {name: 'Landkode', id: 'forretningsadresse', sub:'landkode'},
    {name: 'Land', id: 'forretningsadresse', sub:'land'}
  ];

  return fields;
}

function unwrapObject(item) {
  var out = [];
  var itemFields = getFields();

  itemFields.forEach(function(prop){
    if (prop.sub) {
      if (item[prop.id]) {
        out.push(item[prop.id][prop.sub] || '');
      }
    } else {
      out.push(item[prop.id] || '');
    }
  });

  return out;
}

function writeData(data) {
  var sheet = SpreadsheetApp.getActive();
  data.forEach(function(row){
    sheet.appendRow(row);
  })
}