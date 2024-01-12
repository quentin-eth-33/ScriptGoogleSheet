function downloadPdf(courbeURI, distributionURI, allDataReview) {
  // ID du dossier de destination
  let dossierId = "1zQYZ1rknz4D_C4r70uzJGiSydVgFCYnU";
  
  // Création du document
  let title = "Bilan " + getCurrentDateFormatted();
  let nouveauDoc = DocumentApp.create(title);
  let docId = nouveauDoc.getId();
  // Déplacement du document
  DriveApp.getFileById(nouveauDoc.getId()).moveTo(DriveApp.getFolderById(dossierId));

  // Obtention du corps du document
  let body = nouveauDoc.getBody();

  // Styles
  let titleStyle = {
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_SIZE]: 26,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#351c75',
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
  };

  let subtitleStyle = {
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_SIZE]: 16,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#674ea7',
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  };

  // Ajout du titre
  let h1Title = body.appendParagraph(title);
  h1Title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  h1Title.setAttributes(titleStyle);
  body.appendParagraph('');

  // Sections avec images
  appendSectionWithImage("Courbe Evolution Portefeuille:", courbeURI, body, subtitleStyle, 0.85);
  appendSectionWithImage("Répartition Crypto (Top 10):", distributionURI, body, subtitleStyle, 0.85);

  // Section Statistiques
  let h2TitleStats = body.appendParagraph("Statistique du Mois:");
  h2TitleStats.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h2TitleStats.setAttributes(subtitleStyle);
  body.appendParagraph('');
  displayStatisticsAsTable(allDataReview, body);

  // Section Liste Transactions
  let h2TitleTransac = body.appendParagraph("Liste des Transactions du Mois:");
  h2TitleTransac.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h2TitleTransac.setAttributes(subtitleStyle);
  body.appendParagraph('');
  displayTransactions(allDataReview, body);


  let docText = body.editAsText();
  if (docText.getText().startsWith('\n')) {
    docText.deleteText(0, 0);
  }

  // Sauvegarde et fermeture
  nouveauDoc.saveAndClose();

  // Création du PDF
  let pdfFile = DriveApp.getFileById(nouveauDoc.getId());
  let pdfBlob = pdfFile.getBlob();
  pdfBlob.setName(title + ".pdf");
  DriveApp.getFolderById(dossierId).createFile(pdfBlob);

  DriveApp.getFileById(docId).setTrashed(true);
}


function appendSectionWithImage(title, chart, body, subtitleStyle, reductionRatio) {
  let h2Title = body.appendParagraph(title);
  h2Title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h2Title.setAttributes(subtitleStyle);
  insertImage(chart, body, reductionRatio);
  body.appendParagraph('');
}


function insertImage(chart, body, reductionRatio) {
  let blob = chart.getAs('image/png');
  let image = body.appendImage(blob);

  let originalWidth = image.getWidth();
  let originalHeight = image.getHeight();

  let newWidth = originalWidth * reductionRatio;
  let newHeight = originalHeight * reductionRatio;

  image.setWidth(newWidth);
  image.setHeight(newHeight);
}

function displayStatisticsAsTable(allDataReview, body) {
  let table = body.appendTable();
  table.setBorderWidth(1).setBorderColor('#cccccc'); // Bordure grise pour le tableau

  // Styles pour les cellules de titre
  let titleCellStyle = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#b4a7d6', // Fond bleu
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#FFFFFF', // Texte blanc
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_SIZE]: 12
  };

  // Styles pour les cellules de valeur
  let valueCellStyle = {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000', // Texte noir
    [DocumentApp.Attribute.FONT_SIZE]: 10, // Taille de la police à 12
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#f3f3f3',
  };

  // Tableau de statistiques avec formatage spécifique pour certains champs
  let statistics = [
    { label: "Nombre de Transactions:", value: allDataReview.nbTransactionMonth },
    { label: "Montant Total des Transactions:", value: formatNumberWithDollars(allDataReview.totalAmount) },
    { label: "Nouvelle Crypto:", value: formatList(allDataReview.newCrypto) },
    { label: "Crypto Supprimée:", value: formatList(allDataReview.deleteCrypto) },
    { label: "Montant Nouvel Apport Monétaire:", value: formatNumberWithDollars(allDataReview.amountNewCashIn) },
    { label: "NFT Achetée:", value: formatList(allDataReview.buyNFT) },
    { label: "NFT Vendue:", value: formatList(allDataReview.sellNFT) }
  ];

  // Ajout des lignes au tableau
  statistics.forEach(stat => {
    let row = table.appendTableRow();

    // Appliquer les styles aux cellules de titre
    let titleCell = row.appendTableCell(stat.label);
    titleCell.setAttributes(titleCellStyle);

    // Appliquer les styles aux cellules de valeur
    let valueCell = row.appendTableCell(stat.value.toString());
    valueCell.setAttributes(valueCellStyle);
  });
}

function formatList(cryptoList) {
  // Convertir cryptoList en chaîne de caractères si ce n'est pas déjà le cas
  if (Array.isArray(cryptoList)) {
    cryptoList = cryptoList.join(", ");
  } else if (typeof cryptoList !== 'string') {
    // Si cryptoList n'est ni un tableau ni une chaîne, convertir en chaîne
    cryptoList = String(cryptoList);
  }

  return cryptoList.replace(/,\s*/g, ', ');
}


function displayTransactions(allDataReview, body) {
  let table = body.appendTable();
  table.setBorderWidth(1).setBorderColor('#cccccc'); // Bordure grise pour le tableau

  // Styles pour les en-têtes et les cellules de données
  let headerCellStyle = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#b4a7d6', // Fond bleu
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#FFFFFF', // Texte blanc
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_SIZE]: 10 
  };

  let dataCellStyle = {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.FONT_SIZE]: 8, // Taille de la police à 12
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000',
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#f3f3f3',
  };

  // En-têtes du tableau
  let headers = [
    "Date", "Montant", "Crypto Achetée", "Quantité Achetée", "Prix Moyen Achat",
    "Crypto Vendue", "Quantité Vendue", "Prix Moyen Vente"
  ];

  // Ajout des en-têtes avec styles
  let headerRow = table.appendTableRow();
  headers.forEach(headerText => {
    let cell = headerRow.appendTableCell(headerText);
    cell.setAttributes(headerCellStyle);
  });

  // Ajout des transactions avec styles
  allDataReview.transacInfos.forEach(transac => {
    let row = table.appendTableRow();
    row.appendTableCell(formatDate(transac.date)).setAttributes(dataCellStyle); // Formater la date
    row.appendTableCell(formatNumberWithDollars(transac.amount)).setAttributes(dataCellStyle);
    row.appendTableCell(transac.cryptoBuy).setAttributes(dataCellStyle);
    row.appendTableCell(formatNumber(transac.quantityBuy)).setAttributes(dataCellStyle);
    row.appendTableCell(formatNumberWithDollars(transac.averageBuyPrice)).setAttributes(dataCellStyle);
    row.appendTableCell(transac.cryptoSell).setAttributes(dataCellStyle);
    row.appendTableCell(formatNumber(transac.quantitySell)).setAttributes(dataCellStyle);
    row.appendTableCell(formatNumberWithDollars(transac.averageSellPrice)).setAttributes(dataCellStyle);
  });
}

function formatNumber(averageBuyingPrice) {
  // Convertir en nombre si ce n'est pas déjà un nombre
  let price = parseFloat(averageBuyingPrice);

  // Vérifier si la conversion est réussie
  if (isNaN(price)) {
    return averageBuyingPrice; // Retourner la valeur originale si la conversion échoue
  }

  let formattedNumber;

  if (price < 30 && price > 20) {
    formattedNumber = price.toFixed(1);
  } else if (price <= 20 && price > 1) {
    formattedNumber = price.toFixed(2);
  } else if (price <= 1 && price > 0.09) {
    formattedNumber = price.toFixed(3);
  } else if (price <= 0.09 && price > 0.009) {
    formattedNumber = price.toFixed(4);
  } else if (price <= 0.009 && price > 0.0009) {
    formattedNumber = price.toFixed(5);
  } else if (price <= 0.0009 && price > 0.00009) {
    formattedNumber = price.toFixed(6);
  } else if (price <= 0.00009 && price > 0.000009) {
    formattedNumber = price.toFixed(7);
  } else {
    formattedNumber = price.toFixed(0);
  }

  // Supprimer les zéros inutiles après la partie décimale
  return formattedNumber.replace(/(\.\d*?)0+$/, "$1").replace(/\.$/, "");
}

function formatNumberWithDollars(averageBuyingPrice) {
  let formattedNumber = formatNumber(averageBuyingPrice);
  return formattedNumber + '$';
}

// Fonction pour formater les dates
function formatDate(dateString) {
  let date = new Date(dateString);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}


