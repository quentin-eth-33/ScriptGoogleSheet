function shiftDrawing(action, rowShift){
  const drawings = SHEET.getDrawings();
  let indexDraw;
  for(let i=0; i<drawings.length; i++){
    if(drawings[i].getOnAction() == action){
      indexDraw = i;
      break;
    }
  }
  if(indexDraw){
    let row = drawings[indexDraw].getContainerInfo().getAnchorRow();
    let column = drawings[indexDraw].getContainerInfo().getAnchorColumn();
    console.log("row: "+row)
    console.log("column: "+column)
    drawings[indexDraw].setPosition(row+rowShift, column, 0, 0);


  } else{
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
    console.log("La valeur n'a pas été trouvée sur la feuille.");
    return null;
  }
}

function findCellsWithValueAndBackgroundAllSheet(searchValue, background) {
  let textFinder = SHEET.createTextFinder(searchValue);
  let results = [];

  while (true) {
    let find = textFinder.findNext();
    if (!find) break;

    let row = find.getRow();
    let column = find.getColumn();
    let cell = SHEET.getRange(row, column);

    if (cell.getBackground() === background) {
      results.push({ row: row, column: column });
    }
  }

  if (results.length == 1) {
    return results[0];
  } else {
    console.log("|findCellsWithValueAndBackgroundAllSheet| Erreur, results.length: " + results.length)
    return null;
  }
}


function getAllCrypto(columnCR, rowFirstCrypto) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let choices = [];
  const values = sheet.getRange(rowFirstCrypto, columnCR, sheet.getLastRow(), 1).getValues();
  for (let j = 0; j < values.length; j++) {
    if (values[j][0] == "Total Investi:" || values[j][0] == "") {
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

// Retourne le format adapté au nombre passé en paramètre
function getCellFormatNumber(averageBuyingPrice) {
  var formatNumber = '0';
  if (averageBuyingPrice < 30 && averageBuyingPrice > 20) {
    formatNumber = '0.0';
  }
  else if (averageBuyingPrice <= 20 && averageBuyingPrice > 1) {
    formatNumber = '0.00';
  }
  else if (averageBuyingPrice <= 1 && averageBuyingPrice > 0.09) {
    formatNumber = '0.000';
  }
  else if (averageBuyingPrice <= 0.09 && averageBuyingPrice > 0.009) {
    formatNumber = '0.0000';
  }
  else if (averageBuyingPrice <= 0.009 && averageBuyingPrice > 0.0009) {
    formatNumber = '0.00000';
  }
  else if (averageBuyingPrice <= 0.0009 && averageBuyingPrice > 0.00009) {
    formatNumber = '0.000000'
  }
  else if (averageBuyingPrice <= 0.00009 && averageBuyingPrice > 0.000009) {
    formatNumber = '0.0000000';
  }

  return formatNumber;
}

// Retourne le format adapté au prix passé en paramètre
function getCellFormatNumberDollars(averageBuyingPrice) {
  var formatNumber = '0$';
  if (averageBuyingPrice < 30 && averageBuyingPrice > 20) {
    formatNumber = '0.0$';
  }
  else if (averageBuyingPrice <= 20 && averageBuyingPrice > 1) {
    formatNumber = '0.00$';
  }
  else if (averageBuyingPrice <= 1 && averageBuyingPrice > 0.09) {
    formatNumber = '0.000$';
  }
  else if (averageBuyingPrice <= 0.09 && averageBuyingPrice > 0.009) {
    formatNumber = '0.0000$';
  }
  else if (averageBuyingPrice <= 0.009 && averageBuyingPrice > 0.0009) {
    formatNumber = '0.00000$';
  }
  else if (averageBuyingPrice <= 0.0009 && averageBuyingPrice > 0.00009) {
    formatNumber = '0.000000$'
  }
  else if (averageBuyingPrice <= 0.00009 && averageBuyingPrice > 0.000009) {
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
  //let ui = SpreadsheetApp.getUi();
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
    console.log("L'index de la columnHistoricalDeposit est : " + columnHistoricalDeposit);

    let rowEurosHistoricalDeposit = getRowIndexInColumnWithValue("Euros:", columnHistoricalDeposit)

    if (rowEurosHistoricalDeposit) {
      console.log("L'index de la rowEurosHistoricalDeposit est : " + rowEurosHistoricalDeposit);
      const urlGetEurosPrice = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?id=20641";
      const apiKeyList = getApiKeyCmcList();
      let indexTabApiKey = 0;

      let options = {
        method: "get",
        headers: {
          "X-CMC_PRO_API_KEY": apiKeyList[indexTabApiKey],
        },
        muteHttpExceptions: true // Prevents throwing an exception for non-2xx responses
      };
      // 20641

      let quoteResponse = UrlFetchApp.fetch(urlGetEurosPrice, options);
      let quoteJsonData = quoteResponse.getContentText();

      if (quoteResponse.getResponseCode() === 200) {
        const quoteJson = JSON.parse(quoteJsonData);
        cryptoPrice = quoteJson.data["20641"].quote.USD.price;
        let cell = sheet.getRange(rowEurosHistoricalDeposit, columnHistoricalDeposit + 1);
        cell.setNumberFormat("0.00$");
        cell.setValue(cryptoPrice);
      }

    } else {
      console.log("La valeur Bitcoin n'a pas été trouvée sur la colonne crypto review.");
      //ui.alert('Erreur', "La valeur Bitcoin n'a pas été trouvée sur la colonne crypto review.", ui.ButtonSet.OK);
    }
  } else {
    console.log("La valeur Bilan Crypto n'a pas été trouvée sur la feuille.");
    //ui.alert('Erreur', "La valeur Bilan Crypto n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}
