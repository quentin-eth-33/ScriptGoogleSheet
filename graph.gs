function generatePdf(){
  let hearderCurrentReview = "Bilan Crypto "+getCurrentDateFormatted();
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let graphSheet = getOrCreateSheet(spreadsheet, 'Graphiques');
  if(!findCellWithValueAllSheetName(hearderCurrentReview, SHEET_NAME_DIAGRAM)){
    copyRangeToAnotherSheet();
  }

  let listTotalEvolution = getAllTotalEvolution();
  let dataDiagramDistribution = getDataDistribution();
  let transacInfos =  getTransacInfos();
  let nbTransactionMonth = transacInfos.transactionCount;
  let totalAmount = transacInfos.totalAmount
  let listTransac = transacInfos.listTransac
  let newCrypto = getNewAndDeleteCrypto().newCrypto;
  let deleteCrypto = getNewAndDeleteCrypto().deleteCrypto;
  let buyNFT = getMonthBuyAndSellNFT().buyNFT;
  let sellNFT = getMonthBuyAndSellNFT().sellNFT;
  let amountNewCashIn = getNewTotalCashIn();
  if (listTotalEvolution.length == 0 ||dataDiagramDistribution.length == 0) {
    console.log("|evolutionTotalGraphic| listTotalEvolution.length == 0 ||dataDiagramDistribution.length == 0");
    return;
  }

  let chartData = [["", "Evolution"]];
  for (let i = listTotalEvolution.length - 1; i >= 0; i--) {
    chartData.push([listTotalEvolution[i].date, listTotalEvolution[i].evol]);
  }

  let allData = {
    evolutionCourbe: chartData,
    distributionDiagram: dataDiagramDistribution,
    review: {
      nbTransactionMonth: nbTransactionMonth,
      totalAmount: totalAmount,
      transacInfos: listTransac,
      deleteCrypto: deleteCrypto,
      newCrypto: newCrypto,
      amountNewCashIn: amountNewCashIn,
      buyNFT: buyNFT,
      sellNFT: sellNFT,
    }
  }
  
  createLineChart(graphSheet, chartData);
  createPieChart(graphSheet, dataDiagramDistribution);
  
  let charts = graphSheet.getCharts();
  downloadPdf(charts[0], charts[1], allData.review)

  deleteSheet("Graphiques");
}

function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

function createLineChart(sheet, data) {
  // Insérer les données dans la feuille
  let range = sheet.getRange(1, 1, data.length, 2);
  range.setValues(data);

  // Créer le graphique
  let chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(range)
    .setPosition(1, 3, 0, 0)
    .setOption('hAxis.title', 'Date')      // Légende de l'axe des abscisses
    .setOption('vAxis.title', 'Evolution en %') // Légende de l'axe des ordonnées
    .build();

  sheet.insertChart(chart);
}

function createPieChart(sheet, data) {
  // Insérer les données dans la feuille
  let startRow = sheet.getLastRow() + 2; // Ajouter une marge entre les graphiques
  let range = sheet.getRange(startRow, 1, data.length, 2);
  range.setValues(data);

  // Créer le graphique
  let chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(startRow, 3, 0, 0)
    .setOption('colors', ['#1c91c0', '#43459d', '#8bbc21', '#910000', '#1aadce', '#ff6600', '#5a2fcf', '#ffc0cb', '#008080', '#ffcc00'])
    .setOption('legend', { position: 'bottom', alignment: 'end', textStyle: {color: '#333333', fontSize: 12} }) // Légende en bas à droite
    .setOption('pieHole', 0.4) // Camembert à trou (style donut)
    .setOption('pieSliceText', 'percentage') // Afficher les pourcentages dans les tranches
    .setOption('pieSliceTextStyle', {color: 'white', fontSize: 12}) // Style de texte des pourcentages
    .setOption('chartArea', {left: '10%', top: '10%', width: '80%', height: '70%'}) // Ajustement de la taille du graphique
    .build();

  sheet.insertChart(chart);
}

function copyRangeToAnotherSheet() {

  if (SHEET && SHEET_HISTORIC) {
    const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);
    const NFT_REVIEW = findCellWithValueAllSheet(NR.VALUE_HEADER)

    if (!CRYPTO_REVIEW ||!NFT_REVIEW) {
      console.log("|copyRangeToAnotherSheet| !CRYPTO_REVIEW ||!NFT_REVIEW");
      return;
    }

    let startRow = CRYPTO_REVIEW.row;
    let endRow = getRowIndexInColumnWithValue("Bilan:", CRYPTO_REVIEW.column);
    if (!endRow || !startRow) {
      console.log("|copyRangeToAnotherSheet| !endRow || !startRow");
      return;
    }
    let nbRow = endRow - startRow + 1;

    if (SHEET_HISTORIC.getRange(2, 2).getValue() != "") {
      SHEET_HISTORIC.insertRowsAfter(1, nbRow + 5);
    }

    let sourceRange = SHEET.getRange(startRow, CRYPTO_REVIEW.column, nbRow, 9)
    let targetRange = SHEET_HISTORIC.getRange(2, 2, sourceRange.getNumRows(), sourceRange.getNumColumns());

    sourceRange.copyTo(targetRange, { contentsOnly: true });
    sourceRange.copyTo(targetRange, { formatOnly: true });
    SHEET_HISTORIC.clearConditionalFormatRules();

    let header = "Bilan Crypto " + getCurrentDateFormatted();
    targetRange.getCell(1, 1).setValue(header);

    let protection = targetRange.protect();
    protection.setWarningOnly(true);

    startRow = NFT_REVIEW.row;
    endRow = getRowIndexInColumnWithValue("Bilan:", NFT_REVIEW.column);

    if (!endRow || !startRow) {
      console.log("|copyRangeToAnotherSheet| !endRow || !startRow");
      return;
    }

    sourceRange = SHEET.getRange(startRow, NFT_REVIEW.column, nbRow, NR.NB_FIELDS)
    targetRange = SHEET_HISTORIC.getRange(2, 12, sourceRange.getNumRows(), sourceRange.getNumColumns());

    sourceRange.copyTo(targetRange, { contentsOnly: true });
    sourceRange.copyTo(targetRange, { formatOnly: true });
    SHEET_HISTORIC.clearConditionalFormatRules();

    header = "Bilan NFT " + getCurrentDateFormatted();
    targetRange.getCell(1, 1).setValue(header);

    protection = targetRange.protect();
    protection.setWarningOnly(true);
    
    return (nbRow+5)
  }
}


function getDataDistribution() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_DIAGRAM);
  const CRYPTO_REVIEW = findCellWithValueAllSheet("Bilan Crypto:");

  if (!CRYPTO_REVIEW) {
    console.log("|diagram| !CRYPTO_REVIEW");
    return;
  }

  const startRow = getRowIndexInColumnWithValue("Crypto:", CRYPTO_REVIEW.column) + 1;
  const endRow = getRowIndexInColumnWithValue("Total Investi:", CRYPTO_REVIEW.column) - 1;
  let sumPercentage =0
  if (!endRow || !startRow) {
    console.log("|diagram| !endRow || !startRow");
    return;
  }
  let nbRow = endRow - startRow + 1;

  if (nbRow > 10) {
    nbRow = 10;
  }
  let percentageDataValues = sheet.getRange(5, 10, nbRow, 1).getValues();
  let legendValues = sheet.getRange(5, 2, nbRow, 1).getValues(); // Column containing legends
  let dataDiagramDistribution = [["Crypto","Pourcentage"]]
  for(let i =0; i< percentageDataValues.length; i++){
    dataDiagramDistribution.push([legendValues[i][0], percentageDataValues[i][0]])
    sumPercentage += percentageDataValues[i][0]
  }
  dataDiagramDistribution.push(["Autres", 1-sumPercentage])
  console.log(dataDiagramDistribution)
  return dataDiagramDistribution;
}

function transformDate(text) {
  const monthMap = {
    "Janvier": "01",
    "Février": "02",
    "Mars": "03",
    "Avril": "04",
    "Mai": "05",
    "Juin": "06",
    "Juillet": "07",
    "Août": "08",
    "Septembre": "09",
    "Octobre": "10",
    "Novembre": "11",
    "Décembre": "12",
  };

  const match = text.match(/Bilan Crypto (.+?) (\d{4})/);

  if (match) {
    const monthName = match[1];
    const year = match[2];

    const monthNumber = monthMap[monthName];

    if (monthNumber) {
      return monthNumber + "/" + year.slice(2);
    }
  }

  return null;
}


function getDateEvol(rowIndex, columnIndex){
  let i =0;
  while(rowIndex - i > 0) {
    let cellBackgroundColor = SHEET_HISTORIC.getRange(rowIndex-i, columnIndex+1).getBackground();
    if (cellBackgroundColor  === TH.BG_CRYPTO_NAME) {
      let val = SHEET_HISTORIC.getRange(rowIndex-i, columnIndex+1).getValue();
      return transformDate(val);
    }
    i++;
  }

  return -1;
}

function getAllTotalEvolution() {
  let data = SHEET_HISTORIC.getDataRange().getValues();
  let keyword = "Evolution Totale:";
  let columnIndex = 1; // Colonne ou il y a "Evolution Totale:"
  let valuesArray = [];
  let date;

  let firstIndex = data.findIndex(row => row[columnIndex] === keyword);

  while (firstIndex !== -1) {
    if (firstIndex < data.length - 1) {
      let valueToRetrieve = data[firstIndex][columnIndex + 1] * 100;
      date = getDateEvol(firstIndex, columnIndex);
      valuesArray.push({evol: valueToRetrieve, date: date});
    }
    firstIndex = data.findIndex((row, index) => index > firstIndex && row[columnIndex] === keyword);
  }
  return valuesArray;
}

function getTransacInfos() {
  let listTransac = []
  
  const indexHCT = findCellWithValueAllSheetName(CHT.VALUE_HEADER, SHEET_NAME_TH);
  if (!indexHCT) {
    console.log("|diagram| !indexHCT");
    return;
  }

  let dataRange = SHEET_TRANSACTION_HISTORIC.getRange(indexHCT.row + 3, indexHCT.column, SHEET_TRANSACTION_HISTORIC.getLastRow(), 8);

  let currentDate = new Date();
  let currentMonth = currentDate.getMonth() + 1; // Les mois commencent à partir de zéro, donc nous ajoutons 1.
  let currentYear = currentDate.getFullYear();

  let transactionCount = 0;
  let totalAmount = 0;
  // Parcourez chaque cellule de la plage pour compter les transactions correspondant au mois et à l'année actuels.
  let values = dataRange.getValues();
  for (let i = 0; i < values.length; i++) {
    let cellAmount = values[i][0];
    if (cellAmount == "") {
      break;
    }
    let cellDate = values[i][7];
    cellDate.setDate(cellDate.getDate() + 1)
    if (cellDate instanceof Date) {
      let cellMonth = cellDate.getMonth() + 1;
      let cellYear = cellDate.getFullYear();
      if (cellMonth == currentMonth && cellYear == currentYear) {
        listTransac.push({amount: values[i][0], cryptoBuy: values[i][1], quantityBuy: values[i][2], averageBuyPrice: values[i][3], cryptoSell: values[i][4], quantitySell: values[i][5], averageSellPrice: values[i][6], date: values[i][7]})
        transactionCount++;
        totalAmount += cellAmount;
      }
    }
  }
  return {
    transactionCount: transactionCount,
    totalAmount: totalAmount,
    listTransac: listTransac
  }
  
}


function getCurrentDateFormatted() {
  let moisEnFrancais = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
  ];
  let today = new Date();
  let mois = moisEnFrancais[today.getMonth()];
  let annee = today.getFullYear();
  return mois + " " + annee;
}

function getLastDateFormatted(today) {
  let moisEnFrancais = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
  ];

  let moisIndex = today.getMonth() - 1;
  let annee = today.getFullYear();

  if (moisIndex < 0) {
    moisIndex = 11;  
    annee -= 1;      
  }

  let mois = moisEnFrancais[moisIndex];
  
  return mois + " " + annee;
}


function getNewAndDeleteCrypto() {
  let hearderCurrentReview = "Bilan Crypto "+getCurrentDateFormatted();
  let headerLastReview = "Bilan Crypto "+getLastDateFormatted(new Date());
  console.log("hearderCurrentReview: "+hearderCurrentReview+" | headerLastReview: "+headerLastReview)
  let indexCurrentReview = findCellWithValueAllSheetName(hearderCurrentReview, SHEET_NAME_DIAGRAM);
  let indexLastReview = findCellWithValueAllSheetName(headerLastReview, SHEET_NAME_DIAGRAM);
  if(!indexCurrentReview || !indexLastReview){
    console.log("|getNewAndDeleteCrypto| !indexCurrentReview || !indexLastReview");
    return {
      deleteCrypto: "Aucune",
      newCrypto: "Aucune"
    }
  }
  let valuesCurrentReview = SHEET_HISTORIC.getRange(indexCurrentReview.row+3, indexCurrentReview.column, SHEET_HISTORIC.getLastRow(), 1).getValues();
  let current = [];
  for(let i=0; i<valuesCurrentReview.length; i++){
    if(valuesCurrentReview[i][0] == "Total Investi:"){
      break;
    }
    current.push(valuesCurrentReview[i][0])
  }

  let valuesLastReview = SHEET_HISTORIC.getRange(indexLastReview.row+3, indexLastReview.column, SHEET_HISTORIC.getLastRow(), 1).getValues();
  let last = [];
  for(let i=0; i<valuesLastReview.length; i++){
    if(valuesLastReview[i][0] == "Total Investi:"){
      break;
    }
    last.push(valuesLastReview[i][0])
  }

  let deleteCrypto = last.filter(value => !current.includes(value));
  let newCrypto = current.filter(value => !last.includes(value));

  if(deleteCrypto.length == 0){
    deleteCrypto = "Aucune"
  }
  if(newCrypto.length == 0){
    newCrypto = "Aucune"
  }

  return {
    deleteCrypto: deleteCrypto,
    newCrypto: newCrypto
  }
}

function getNewTotalCashIn() {
  const indexHTI = findCellWithValueAllSheet(HTI.VALUE_HEADER);
  if (!indexHTI) {
    console.log("|getNewTotalCashIn| !indexHTI");
    return;
  }

  let dataRange = SHEET.getRange(indexHTI.row + 3, indexHTI.column, SHEET.getLastRow(), 2);

  let currentDate = new Date();
  let currentMonth = currentDate.getMonth() + 1; // Les mois commencent à partir de zéro, donc nous ajoutons 1.
  let currentYear = currentDate.getFullYear();

  let totalAmount = 0;
  // Parcourez chaque cellule de la plage pour compter les transactions correspondant au mois et à l'année actuels.
  let values = dataRange.getValues();
  for (let i = 0; i < values.length; i++) {
    let cellDate = values[i][0];
    let cellAmount = values[i][1];
    if (cellAmount == "") {
      break;
    }
    else if (cellDate instanceof Date) {
      let cellMonth = cellDate.getMonth() + 1;
      let cellYear = cellDate.getFullYear();
      if (cellMonth == currentMonth && cellYear == currentYear) {
        totalAmount += cellAmount;
      }
    }
  }
  return totalAmount;
}

function getMonthBuyAndSellNFT() {
  let hearderCurrentReview = "Bilan NFT "+getCurrentDateFormatted();
  let headerLastReview = "Bilan NFT "+getLastDateFormatted(new Date());

  console.log("hearderCurrentReview: "+hearderCurrentReview+" | headerLastReview: "+headerLastReview)
  let indexCurrentReview = findCellWithValueAllSheetName(hearderCurrentReview, SHEET_NAME_DIAGRAM);
  let indexLastReview = findCellWithValueAllSheetName(headerLastReview, SHEET_NAME_DIAGRAM);

  if(!indexCurrentReview || !indexLastReview){
    console.log("|getMonthBuyAndSellNFT| !indexCurrentReview || !indexLastReview");
    return {
      sellNFT: "Aucune",
      buyNFT: "Aucune"
    }
  }
  let valuesCurrentReview = SHEET_HISTORIC.getRange(indexCurrentReview.row+3, indexCurrentReview.column, SHEET_HISTORIC.getLastRow(), 1).getValues();
  let current = [];
  for(let i=0; i<valuesCurrentReview.length; i++){
    if(valuesCurrentReview[i][0] == "Total Investi:"){
      break;
    }
    current.push(valuesCurrentReview[i][0])
  }

  let valuesLastReview = SHEET_HISTORIC.getRange(indexLastReview.row+3, indexLastReview.column, SHEET_HISTORIC.getLastRow(), 1).getValues();
  let last = [];
  for(let i=0; i<valuesLastReview.length; i++){
    if(valuesLastReview[i][0] == "Total Investi:"){
      break;
    }
    last.push(valuesLastReview[i][0])
  }

  let sellNFT = last.filter(value => !current.includes(value));
  let buyNFT = current.filter(value => !last.includes(value));

  if(sellNFT.length == 0){
    sellNFT = "Aucune"
  }
  if(buyNFT.length == 0){
    buyNFT = "Aucune"
  }

  return {
    sellNFT: sellNFT,
    buyNFT: buyNFT
  }
}

function showReview(){
  let ui = SpreadsheetApp.getUi();
  let nbRow = copyRangeToAnotherSheet();
  let listTotalEvolution = getAllTotalEvolution();
  let dataDiagramDistribution = getDataDistribution();
  let transacInfos =  getTransacInfos();
  let nbTransactionMonth = transacInfos.transactionCount;
  let totalAmount = transacInfos.totalAmount
  let listTransac = transacInfos.listTransac
  let newCrypto = getNewAndDeleteCrypto().newCrypto;
  let deleteCrypto = getNewAndDeleteCrypto().deleteCrypto;
  let buyNFT = getMonthBuyAndSellNFT().buyNFT;
  let sellNFT = getMonthBuyAndSellNFT().sellNFT;
  let amountNewCashIn = getNewTotalCashIn();
  if (listTotalEvolution.length == 0 ||dataDiagramDistribution.length == 0) {
    console.log("|evolutionTotalGraphic| listTotalEvolution.length == 0 ||dataDiagramDistribution.length == 0");
    SpreadsheetApp.getUi().alert('Erreur', "|evolutionTotalGraphic| listTotalEvolution.length == 0 ||dataDiagramDistribution.length == 0", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let chartData = [["Date", "Evolution"]];
  for (let i = listTotalEvolution.length - 1; i >= 0; i--) {
    chartData.push([listTotalEvolution[i].date, listTotalEvolution[i].evol]);
  }

  let allData = {
    evolutionCourbe: chartData,
    distributionDiagram: dataDiagramDistribution,
    review: {
      nbTransactionMonth: nbTransactionMonth,
      totalAmount: totalAmount,
      transacInfos: listTransac,
      deleteCrypto: deleteCrypto,
      newCrypto: newCrypto,
      amountNewCashIn: amountNewCashIn,
      buyNFT: buyNFT,
      sellNFT: sellNFT,
    }
  }

  SHEET_HISTORIC.deleteRows(2,nbRow)
  let htmlOutput = HtmlService.createHtmlOutputFromFile('pageDiagram')
    .setWidth(900)
    .setHeight(900);
  htmlOutput.append('<script>let allData = ' + JSON.stringify(allData) + ';</script>');
  ui.showModalDialog(htmlOutput, "Bilan "+getCurrentDateFormatted()+":");
}
