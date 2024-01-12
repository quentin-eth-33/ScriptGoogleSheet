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
  appendSectionWithImageV2("Courbe Evolution Portefeuille:", courbeURI, body, subtitleStyle, 0.85);
  appendSectionWithImageV2("Répartition Crypto (Top 10):", distributionURI, body, subtitleStyle, 0.85);

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


function appendSectionWithImageV2(title, chart, body, subtitleStyle, reductionRatio) {
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


