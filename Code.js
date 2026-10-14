function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Team Wiki (Tettra L2)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper function to include HTML/CSS/JS files into main template
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Safely fetch all wiki articles
function getArticles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Wiki');
    
    // Auto-create the sheet with headers if the user deleted or misspelled it
    if (!sheet) {
      sheet = ss.insertSheet('Wiki');
      sheet.appendRow(['ID', 'Title', 'Category', 'Content']);
    }
    
    const data = sheet.getDataRange().getValues();
    const articles = [];
    
    if (data.length <= 1) return []; // Only headers present
    
    for (let i = 1; i < data.length; i++) {
      // Direct variable mapping to prevent array reference errors
      const rowId = data[i][0] ? data[i][0].toString() : '';
      const rowTitle = data[i][1] ? data[i][1].toString() : '';
      const rowCategory = data[i][2] ? data[i][2].toString() : '';
      const rowContent = data[i][3] ? data[i][3].toString() : '';
      
      if (rowId || rowTitle) {
        articles.push({
          id: rowId,
          title: rowTitle,
          category: rowCategory,
          content: rowContent
        });
      }
    }
    return articles;
  } catch (error) {
    Logger.log("Error in getArticles: " + error.toString());
    return { error: error.toString() };
  }
}

// Save or Update an article
function saveArticle(id, title, category, content) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Wiki');
    const data = sheet.getDataRange().getValues();
    
    if (!id) {
      const newId = 'art_' + new Date().getTime();
      sheet.appendRow([newId, title, category, content]);
      return { success: true, message: 'Article created successfully!', generatedId: newId };
    } else {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === id.toString()) {
          sheet.getRange(i + 1, 2, 1, 3).setValues([[title, category, content]]);
          return { success: true, message: 'Article updated successfully!', generatedId: id };
        }
      }
    }
    return { success: false, message: 'Article ID matching failed.' };
  } catch (error) {
    return { success: false, message: 'Backend Error: ' + error.toString() };
  }
}
