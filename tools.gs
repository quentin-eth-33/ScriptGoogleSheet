function deleteSheet(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet != null) {
    spreadsheet.deleteSheet(sheet);
  } else {
    console.log("La feuille n'a pas été trouvée.");
  }
}


function shiftDrawing(action, rowShift) {
  const drawings = SHEET.getDrawings();
  let indexDraw;
  for (let i = 0; i < drawings.length; i++) {
    if (drawings[i].getOnAction() == action) {
      indexDraw = i;
      break;
    }
  }
  if (indexDraw) {
    let row = drawings[indexDraw].getContainerInfo().getAnchorRow();
    let column = drawings[indexDraw].getContainerInfo().getAnchorColumn();
    console.log("row: " + row)
    console.log("column: " + column)
    drawings[indexDraw].setPosition(row + rowShift, column, 0, 0);


  } else {
    console.log("|shiftDrawing| Dessin pas trouvé");
  }
}


// Renvoie l'index de la ligne/colonne de la valeur recherchée. La recherche s'effectue sur toute la feuille.
function findCellWithValueAllSheet(searchValue) {
  let find = SHEET.createTextFinder(searchValue).findNext();
  if (find) {
    let row = find.getRow();
    let column = find.getColumn();
    let results = {
      row: row,
      column: column
    };
    return results;
  } else {
    console.log("|findCellWithValueAllSheet| La valeur n'a pas été trouvée sur la feuille.");
    return null;
  }
}

function findCellWithValueAllSheetName(searchValue, sheetName) {
  const sheetNameTemp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let find = sheetNameTemp.createTextFinder(searchValue).findNext();
  if (find) {
    let row = find.getRow();
    let column = find.getColumn();
    let results = {
      row: row,
      column: column
    };
    return results;
  } else {
    console.log("|findCellWithValueAllSheetName| La valeur n'a pas été trouvée sur la feuille.");
    return null;
  }
}

function findCellsWithValueAndBackgroundAllSheet(sheetName, searchValue, background) {
  const SHEET_TH = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let textFinder = SHEET_TH.createTextFinder(searchValue);
  let results = [];

  while (true) {
    let find = textFinder.findNext();
    if (!find) break;

    let row = find.getRow();
    let column = find.getColumn();
    let cell = SHEET_TH.getRange(row, column);

    if (cell.getBackground() === background) {
      results.push({ row: row, column: column });
    }
  }

  if (results.length > 0) {
    return results[0];
  } else {
    console.log("|findCellsWithValueAndBackgroundAllSheet| Erreur, results.length: " + results.length)
    return null;
  }
}

// Fonction de comparaison pour le tri décroissant en fonction de la valeur "value"
function descendingComparisonFunction(a, b) {
  return b.value - a.value;
}


function getAllCrypto(columnCR, rowFirstCrypto) {
  let choices = [];
  console.log("|getAllCrypto| rowFirstCrypto: "+rowFirstCrypto)
  const values = SHEET.getRange(rowFirstCrypto, columnCR, SHEET.getLastRow(), 1).getValues();
  for (let j = 0; j < values.length; j++) {
    if (values[j][0] == CR.VALUE_AFTER_LAST_CRYPTO || values[j][0] == "") {
      break;
    }
    else {
      choices.push(values[j][0])
    }
  }
  return choices;
}

// Renvoie l'index de la ligne contenant la valeur "searchValue" sur la colonne "indexColumn"
function getRowIndexInColumnWithValue(searchValue, indexColumn) {
  var columnRange = SHEET.getRange(1, indexColumn, SHEET.getLastRow(), 1);
  var columnValue = columnRange.getValues();
  for (var i = 0; i < columnValue.length; i++) {
    if (columnValue[i][0] === searchValue) {
      var indexRow = i + 1;
      return indexRow;
    }
  }
  console.log("|getRowIndexInColumnWithValue| Aucune ligne ne contient la valeur: " + searchValue + " recherchée sur la colonne: " + indexColumn);
  return null;
}

// Renvoie l'index de la ligne contenant la valeur "searchValue" sur la colonne "indexColumn"
function getRowIndexInColumnWithValueWithSheet(sheetName, searchValue, indexColumn) {
  const SHEET_TH = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var columnRange = SHEET_TH.getRange(1, indexColumn, SHEET_TH.getLastRow(), 1);
  var columnValue = columnRange.getValues();
  for (var i = 0; i < columnValue.length; i++) {
    if (columnValue[i][0] === searchValue) {
      var indexRow = i + 1;
      return indexRow;
    }
  }
  console.log("|getRowIndexInColumnWithValueWithSheet| Aucune ligne ne contient la valeur: " + searchValue + " recherchée sur la colonne: " + indexColumn);
  return null;
}


// Renvoie la lettre de la colonne à partir de son index, ex: si l'index est 5 elle renvoie "E"
function getColumnHeader(columnIndex) {
  let columnHeader = "";
  let tempIndex = columnIndex;

  while (tempIndex > 0) {
    let remainder = (tempIndex - 1) % 26;
    let charCode = 65 + remainder;
    columnHeader = String.fromCharCode(charCode) + columnHeader;
    tempIndex = Math.floor((tempIndex - 1) / 26);
  }
  return columnHeader;
}

// Défini le format des cellules bleues
function setBleueTexte(indexRowNewCell, indexColumnNewCell, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  cell = sheet.getRange(indexRowNewCell, indexColumnNewCell);
  // police #4285f4
  // background: #cfe2f3
  cell.setBackground("#cfe2f3");
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("left");
  cell.setFontSize(8);
  cell.setFontColor("#4285f4");
  cell.setValue(value);
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
}

// Permet de fusionner des celulles à partir de range
function mergeCells(startRange, endRange) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeToMerge = sheet.getRange(startRange.getRow(), startRange.getColumn(), endRange.getRow() - startRange.getRow() + 1, endRange.getColumn() - startRange.getColumn() + 1);
  rangeToMerge.merge(); // Fusionner les cellules
}

// Obtenir la référence du background
function getBackgroundColorReference(indexRow, indexColumn) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange(indexRow, indexColumn);
  var backgroundColor = cell.getBackground();
  console.log("Référence du fond de la cellule: " + backgroundColor);
}

// Obtenir la référence de la police
function getFontColorReference(indexRow, indexColumn) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange(indexRow, indexColumn);
  var fontColor = cell.getFontColor();
  console.log("Couleur de la police de la cellule: " + fontColor);
}

function color(){
  getFontColorReference(1,1);
  getBackgroundColorReference(1,1)
}

// Renvoie l'index de la dernière colonne de la cellule fusionnée passé en paramètre
function getMergedCellColumnsEnd(rowIndex, columnIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange(rowIndex, columnIndex);
  var mergedRanges = range.getMergedRanges();

  if (mergedRanges.length > 0) {
    var mergedRange = mergedRanges[0];
    var startColumn = mergedRange.getColumn();
    var endColumn = startColumn + mergedRange.getWidth() - 1;

    console.log("Colonne de début : " + startColumn);
    console.log("Colonne de fin : " + endColumn);
  }
  return endColumn;
}

function getCellFormatNumber(averageBuyingPrice) {
  // Convertir averageBuyingPrice en nombre si ce n'est pas déjà un nombre
  var price = (typeof averageBuyingPrice === 'string') ? parseFloat(averageBuyingPrice) : averageBuyingPrice;

  var formatNumber = '0';

  // Vérifie si price est un entier
  if (Number.isInteger(price)) {
    formatNumber = '0';
  } else if (price < 30 && price > 20) {
    formatNumber = '0.0';
  } else if (price <= 20 && price > 1) {
    formatNumber = '0.00';
  } else if (price <= 1 && price > 0.09) {
    formatNumber = '0.000';
  } else if (price <= 0.09 && price > 0.009) {
    formatNumber = '0.0000';
  } else if (price <= 0.009 && price > 0.0009) {
    formatNumber = '0.00000';
  } else if (price <= 0.0009 && price > 0.00009) {
    formatNumber = '0.000000';
  } else if (price <= 0.00009 && price > 0.000009) {
    formatNumber = '0.0000000';
  }

  return formatNumber;
}



function getCellFormatNumberDollars(averageBuyingPrice) {
  // Convertir averageBuyingPrice en nombre si ce n'est pas déjà un nombre
  var price = (typeof averageBuyingPrice === 'string') ? parseFloat(averageBuyingPrice) : averageBuyingPrice;

  var formatNumber ='0$';

  // Vérifie si price est un entier
  if (Number.isInteger(price)) {
    formatNumber = '0$';
  } else if (price < 30 && price > 20) {
    formatNumber = '0.0$';
  } else if (price <= 20 && price > 1) {
    formatNumber = '0.00$';
  } else if (price <= 1 && price > 0.09) {
    formatNumber = '0.000$';
  } else if (price <= 0.09 && price > 0.009) {
    formatNumber = '0.0000$';
  } else if (price <= 0.009 && price > 0.0009) {
    formatNumber = '0.00000$';
  } else if (price <= 0.0009 && price > 0.00009) {
    formatNumber = '0.000000$';
  } else if (price <= 0.00009 && price > 0.000009) {
    formatNumber = '0.0000000$';
  }

  return formatNumber;
}

// Initialisation d'une variable globale
function initGlobalVariable(globalVariableName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const globalVariable = "";
  scriptProperties.setProperty(globalVariableName, globalVariable);
}


function getApiKeyCmcList() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const API_KEY = scriptProperties.getProperty('API-KEY');
  return API_KEY.split(";");
}

function getFormatCell(value) {
  console.log("Tools: value= ", value);
  console.log("Tools: value.indexOf('.')= ", value.indexOf('.'));
  if (value.indexOf('.') !== -1) {
    value = value.replace('.', ',');
  }
  return value;
}

function getFormatCalculationScript(value) {
  if (value.indexOf(',') !== -1) {
    value = value.replace(',', '.');
  }
  return value;
}

function getCellDateFormat(date) {
  let dateComponents = date.split("-");
  let formattedDate = dateComponents[2] + "/" + dateComponents[1] + "/" + dateComponents[0];
  return formattedDate;
}

function refreshEurosPrice() {
  let ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  let mapIdCryptoGlobalString = scriptProperties.getProperty('mapIdCryptoGlobal');

  // Convertir la chaîne en map
  const mapIdCryptoGlobal = {};
  const keyValuePairs = mapIdCryptoGlobalString.split(";");
  keyValuePairs.forEach(keyValuePair => {
    const [key, value] = keyValuePair.split(":");
    mapIdCryptoGlobal[key] = value;
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let historicalDeposit = findCellWithValueAllSheet("Historique Total Investi:");

  if (historicalDeposit) {
    let columnHistoricalDeposit = historicalDeposit.column;
    let rowEurosHistoricalDeposit = getRowIndexInColumnWithValue(HTI.VALUE_ROW_EUROS, columnHistoricalDeposit)

    if (rowEurosHistoricalDeposit) {
      let eurosPrice = getEurosToUsdPrice();
      if (eurosPrice) {
        let cell = sheet.getRange(rowEurosHistoricalDeposit, columnHistoricalDeposit + 1);
        cell.setNumberFormat("0.00$");
        cell.setValue(eurosPrice);
      }
      else{
        console.log("|refreshEurosPrice| getEurosToUsdPrice fail");
      }
    } else {
      console.log("La valeur Bitcoin n'a pas été trouvée sur la colonne crypto review.");
      ui.alert('Erreur', "La valeur Bitcoin n'a pas été trouvée sur la colonne crypto review.", ui.ButtonSet.OK);
    }
  } else {
    console.log("La valeur Bilan Crypto n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "La valeur Bilan Crypto n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}


function getEurosToUsdPrice() {
  let ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  let mapIdCryptoGlobalString = scriptProperties.getProperty('mapIdCryptoGlobal');

  // Convertir la chaîne en map
  const mapIdCryptoGlobal = {};
  const keyValuePairs = mapIdCryptoGlobalString.split(";");
  keyValuePairs.forEach(keyValuePair => {
    const [key, value] = keyValuePair.split(":");
    mapIdCryptoGlobal[key] = value;
  });

  const urlGetEurosPrice = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?id=20641";
  const apiKeyList = getApiKeyCmcList();
  let indexTabApiKey = 0;

  let options = {
    method: "get",
    headers: {
      "X-CMC_PRO_API_KEY": apiKeyList[indexTabApiKey],
    },
    muteHttpExceptions: true 
  };
  // 20641

  let quoteResponse = UrlFetchApp.fetch(urlGetEurosPrice, options);
  let quoteJsonData = quoteResponse.getContentText();

  if (quoteResponse.getResponseCode() === 200) {
    const quoteJson = JSON.parse(quoteJsonData);
    let cryptoPrice = quoteJson.data["20641"].quote.USD.price;
    return cryptoPrice;
  }
  return null;
}

function getCryptoPriceWithID(idCrypto) {


  const urlGetEurosPrice = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?id="+idCrypto
  const apiKeyList = getApiKeyCmcList();
  let indexTabApiKey = 0;

  let options = {
    method: "get",
    headers: {
      "X-CMC_PRO_API_KEY": apiKeyList[indexTabApiKey],
    },
    muteHttpExceptions: true 
  };

  let quoteResponse = UrlFetchApp.fetch(urlGetEurosPrice, options);
  let quoteJsonData = quoteResponse.getContentText();

  if (quoteResponse.getResponseCode() === 200) {
    const quoteJson = JSON.parse(quoteJsonData);
    let cryptoPrice = quoteJson.data[idCrypto].quote.USD.price;
    return cryptoPrice;
  }
  return null;
}


