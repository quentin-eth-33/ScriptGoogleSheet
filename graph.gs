
function graph(){
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
