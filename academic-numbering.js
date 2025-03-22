// Written by Christina Natasha from Bret Jordan's code
// https://github.com/kurisu-na/academic-numbering
// Version 1.0.0
// Last updated 2025-03-14
// Apache 2.0 license

/**
  * @OnlyCurrentDoc
 */

function onOpen(e) {
  var ui = DocumentApp.getUi();
  ui.createAddonMenu()
  .addSubMenu(ui.createMenu('Heading Numbers (up to Heading 4)')
              .addItem('1.2.3', 'addSimpleHeadingNumbers')
              .addItem('1.2.3.', 'addSimpleHeadingNumbersDot')
             )
  .addSubMenu(ui.createMenu('Figure Numbers (Heading 5)')
              .addItem('Figure 1', 'addFigNum')
              .addItem('Figure 1.', 'addFigNumDot')
             )
  .addSubMenu(ui.createMenu('Table Numbers (Heading 6)')
              .addItem('Table 1', 'addTblNum')
              .addItem('Table 1.', 'addTblNumDot')
             )
  .addToUi();
}


// Run on install
function onInstall(e) {
  onOpen(e);
}


// Menu Functions

// 1.2.3
function addSimpleHeadingNumbers(){
  headingNumbers(true, 1, false);
}

// 1.2.3.
function addSimpleHeadingNumbersDot(){
  headingNumbers(true, 1, true);
}

// Figure 1
function addFigNum(){
  figNum(true, false);
}

// Figure 1.
function addFigNumDot(){
  figNum(true, true);
}

// Table 1
function addTblNum(){
  tblNum(true, false);
}

// Table 1.
function addTblNumDot(){
  tblNum(true, true);
}


// Main functions
function figNum(add, enddot){
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  var paragraphs = document.getParagraphs();
  
  // Stores the count of figures
  var figureCount = 0;

  // Loop through all of the paragraphs that we have found
  for (var i in paragraphs) {
    var element = paragraphs[i];
    var type = element.getHeading();
    
    // exclude everything but Heading 5
    if (type != DocumentApp.ParagraphHeading.HEADING5) {
      continue;
    }

    if(add == true) {
      // The actual numbering to insert
      var figureNum = 'Figure ';

      figureCount++;

      if(enddot == true){
        figureNum += figureCount + '.';
      }
      else {
        figureNum += figureCount;
      }      
    }


    // If there is an existing header number with out without a markdown hash, remove it.

    // The first regex looks for an existing markdown header hash mark and heading number.
    element.replaceText('^(Figure+\\s)?([(0-9)*\\-)(\\.[(0-9)*\\-])*\\.?\\s', '');

    // The second regex looks for an existing markdown header hash mark and heading letter.
    // I am doing it this way so in the future I can address the letters in a different way.
    element.replaceText('^(Figure+\\s)?([a-zA-Z\\-])(\\.[a-zA-Z\\-])*\\.?\\s', '');


    // Insert the actual text
    element.insertText(0, figureNum + ' ');
    element.setHeading(DocumentApp.ParagraphHeading.HEADING5);  
  }
}

function tblNum(add, enddot){
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  var paragraphs = document.getParagraphs();
  
  // Stores the count of figures
  var tableCount = 0;

  // Loop through all of the paragraphs that we have found
  for (var i in paragraphs) {
    var element = paragraphs[i];
    var type = element.getHeading();
    
    // exclude everything but Heading 6
    if (type != DocumentApp.ParagraphHeading.HEADING6) {
      continue;
    }
    
    if(add == true) {
      // The actual numbering to insert
      var tableNum = 'Table ';

      tableCount++;

      if(enddot == true){
        tableNum += tableCount + '.';
      }
      else {
        tableNum += tableCount;
      }
    }

    // If there is an existing header number with out without a markdown hash, remove it.

    // The first regex looks for an existing markdown header hash mark and heading number.
    element.replaceText('^(Table+\\s)?([(0-9)*\\-)(\\.[(0-9)*\\-])*\\.?\\s', '');

    // The second regex looks for an existing markdown header hash mark and heading letter.
    // I am doing it this way so in the future I can address the letters in a different way.
    element.replaceText('^(Table+\\s)?([a-zA-Z\\-])(\\.[a-zA-Z\\-])*\\.?\\s', '');

    // Insert the actual text
    element.insertText(0, tableNum + ' ');
    element.setHeading(DocumentApp.ParagraphHeading.HEADING6);
  }
}

function headingNumbers(add, headertype, enddot){
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  var paragraphs = document.getParagraphs();
  
  // This array of numbers will increment as we encounter new headings and thus keep 
  // the state of where we are at.  This is how the code knows which number is next.
  var numbers = [0,0,0,0,0,0,0];
  
  // Loop through all of the paragraphs that we have found
  for (var i in paragraphs) {
    var element = paragraphs[i];
    var text = element.getText()+'';
    var type = element.getHeading()+'';
    
    // exclude everything but headings
    if (!type.match(/HEADING\d/)) {
      continue;
    }
    
    // exclude empty headings (e.g. page breaks generate these)
    if( text.match(/^\s*$/)){
      continue;
    }

    if (add == true) {
      // Get the current header level (e.g., 1, 2, 3, 4, 5, 6)
      var headerlevel = new RegExp(/HEADING(\d)/).exec(type)[1];
      
      // This string will hold the actual header numbering value, be it integers or letters
      var numbering = '';
      
      // Increment the heading number in the array. This will enable us to know 
      // where we are at in the numbering.
      numbers[headerlevel]++;
      
      // Loop through all 6 possible levels and concatinate them in to a single string
      // Check for the type level and as needed, convert the numbers array in to strings
      // either numerical strings or letters.
      for (var level = 1; level <= 6; level++) {
        if (level < headerlevel) {
          if (headertype == 1) {
            numbering += numbers[level] + '.';
          }
        } else if (level == headerlevel) {
          if (enddot) {
            if (headertype == 1) {
              numbering += numbers[level] + '.';
            }
          } else {
            if (headertype == 1) {
              numbering += numbers[level];
            }
          }
        } else {
          numbers[level] = 0;
        }
      }

      // If there is an existing header number with out without a markdown hash, remove it.

      // The first regex looks for an existing markdown header hash mark and heading number.
      element.replaceText('^(#+\\s)?([(0-9)*\\-)(\\.[(0-9)*\\-])*\\.?\\s', '');

      // The second regex looks for an existing markdown header hash mark and heading letter.
      // I am doing it this way so in the future I can address the letters in a different way.
      element.replaceText('^(#+\\s)?([a-zA-Z\\-])(\\.[a-zA-Z\\-])*\\.?\\s', '');
       
      // Check each heading level and set the right numbers and apply the right style 
      if (headerlevel == 1) {
        element.insertText(0, numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      } else if (headerlevel == 2) {
        element.insertText(0, numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      } else if (headerlevel == 3) {
        element.insertText(0, numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      } else if (headerlevel == 4) {
        element.insertText(0, numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING4);
      }
    } else {
      // Clear current header numbers
      // If there is an existing header number with out without a markdown hash, remove it.

      // The first regex looks for an existing markdown header hash mark and heading number.
      element.replaceText('^(#+\\s)?([(0-9)*\\-)(\\.[(0-9)*\\-])*\\.?\\s', '');

      // The second regex looks for an existing markdown header hash mark and heading letter.
      // I am doing it this way so in the future I can address the letters in a different way.
      element.replaceText('^(#+\\s)?([a-zA-Z\\-])(\\.[a-zA-Z\\-])*\\.?\\s', '');
    }
  }
}
