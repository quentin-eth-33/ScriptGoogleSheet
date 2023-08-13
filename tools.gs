// Renvoie l'index de la ligne/colonne de la valeur recherchée. La recherche s'effectue sur toute la feuille.
function findCellWithValueAllSheet(searchValue) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let find = sheet.createTextFinder(searchValue).findNext();
  if (find) {
    let row = find.getRow();
    let column = find.getColumn();
    Logger.log("L'index de la ligne est : " + row);
    Logger.log("L'index de la colonne est : " + column);
    let results = {
      row: row,
      column: column
    };
    return results;
  } else {
    Logger.log("La valeur n'a pas été trouvée sur la feuille.");
    return null;
  }
}

// Renvoie l'index de la ligne contenant la valeur "searchValue" sur la colonne "indexColumn"
function getRowIndexInColumnWithValue(searchValue, indexColumn) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnRange = sheet.getRange(1, indexColumn, sheet.getLastRow(), 1);
  var columnValue = columnRange.getValues();
  for (var i = 0; i < columnValue.length; i++) {
    if (columnValue[i][0] === searchValue) {
      var indexRow = i + 1; // Ajouter 1 car les indices de ligne commencent à partir de 1
      Logger.log("L'index de la ligne contenant la valeur cherchée est : " + indexRow);
      return indexRow;
    }
  }
  Logger.log("Aucune ligne ne contient la valeur recherchée sur la colonne: " + indexColumn);
  return null; // Retourne -1 si aucune occurrence n'est trouvée.
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
  console.log("startRange.getRow(): " + startRange.getRow())
  var rangeToMerge = sheet.getRange(startRange.getRow(), startRange.getColumn(), endRange.getRow() - startRange.getRow() + 1, endRange.getColumn() - startRange.getColumn() + 1);
  rangeToMerge.merge(); // Fusionner les cellules
}

// Obtenir la référence du background
function getBackgroundColorReference(indexRow, indexColumn) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange(indexRow, indexColumn);
  var backgroundColor = cell.getBackground();
  Logger.log("Référence du fond de la cellule: " + backgroundColor);
}

// Obtenir la référence de la police
function getFontColorReference(indexRow, indexColumn) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange(indexRow, indexColumn);
  var fontColor = cell.getFontColor();
  Logger.log("Couleur de la police de la cellule: " + fontColor);
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

    Logger.log("Colonne de début : " + startColumn);
    Logger.log("Colonne de fin : " + endColumn);
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
