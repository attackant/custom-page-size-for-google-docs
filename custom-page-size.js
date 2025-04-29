/**
 * Custom Page Size for Google Docs (with KDP presets)
 * 
 * This script adds custom page size functionality to Google Docs with KDP presets,
 * allowing authors to format their documents according to Kindle Direct Publishing
 * book size requirements.
 * 
 * Features:
 * - Standard KDP paperback and hardcover sizes
 * - Common paper sizes (Letter, A4, etc.)
 * - Custom sizing with KDP-compliant validation
 * - Paper type selection (white/cream)
 * - Ink type selection (black/color)
 * - Margin presets and custom margins
 * 
 * @license GPL-3.0
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, version 3 of the License.
 * 
 * @version 1.1.0
 * @author Damian Taggart
 * @link https://github.com/attackant/custom-page-size-for-google-docs
 */

// ========================
// CONSTANTS AND DATA
// ========================

// Conversion constants
const POINTS_PER_INCH = 72;

// Color constants
const COLOR_WHITE = '#FFFFFF';
const COLOR_CREAM = '#F8F3E6';

// KDP size constraints
const MIN_WIDTH = 4;      // Minimum width in inches
const MAX_WIDTH = 8.5;    // Maximum width in inches
const MIN_HEIGHT = 6;     // Minimum height in inches
const MAX_HEIGHT = 11.69; // Maximum height in inches

// Define KDP book sizes (in inches)
const KDP_SIZES = {
  // Paperback sizes
  'Paperback - 5 x 8': { width: 5, height: 8, type: 'paperback' },
  'Paperback - 5.06 x 7.81': { width: 5.06, height: 7.81, type: 'paperback' },
  'Paperback - 5.25 x 8': { width: 5.25, height: 8, type: 'paperback' },
  'Paperback - 5.5 x 8.5': { width: 5.5, height: 8.5, type: 'paperback' },
  'Paperback - 6 x 9': { width: 6, height: 9, type: 'paperback' },
  'Paperback - 6.14 x 9.21': { width: 6.14, height: 9.21, type: 'paperback' },
  'Paperback - 6.69 x 9.61': { width: 6.69, height: 9.61, type: 'paperback' },
  'Paperback - 7 x 10': { width: 7, height: 10, type: 'paperback' },
  'Paperback - 7.44 x 9.69': { width: 7.44, height: 9.69, type: 'paperback' },
  'Paperback - 7.5 x 9.25': { width: 7.5, height: 9.25, type: 'paperback' },
  'Paperback - 8 x 10': { width: 8, height: 10, type: 'paperback' },
  'Paperback - 8.25 x 6': { width: 8.25, height: 6, type: 'paperback' },
  'Paperback - 8.25 x 8.25': { width: 8.25, height: 8.25, type: 'paperback' },
  'Paperback - 8.5 x 8.5': { width: 8.5, height: 8.5, type: 'paperback' },
  'Paperback - 8.5 x 11': { width: 8.5, height: 11, type: 'paperback' },
  
  // Hardcover sizes
  'Hardcover - 5 x 8': { width: 5, height: 8, type: 'hardcover' },
  'Hardcover - 5.5 x 8.5': { width: 5.5, height: 8.5, type: 'hardcover' },
  'Hardcover - 6 x 9': { width: 6, height: 9, type: 'hardcover' },
  'Hardcover - 7 x 10': { width: 7, height: 10, type: 'hardcover' },
  'Hardcover - 8 x 10': { width: 8, height: 10, type: 'hardcover' },
  'Hardcover - 8.25 x 8.25': { width: 8.25, height: 8.25, type: 'hardcover' },
  'Hardcover - 8.5 x 8.5': { width: 8.5, height: 8.5, type: 'hardcover' },
  'Hardcover - 8.5 x 11': { width: 8.5, height: 11, type: 'hardcover' },
};

// Common non-KDP paper sizes
const COMMON_SIZES = {
  'Letter': { width: 8.5, height: 11 },
  'Legal': { width: 8.5, height: 14 },
  'Tabloid': { width: 11, height: 17 },
  'A4': { width: 8.27, height: 11.69 },
  'A5': { width: 5.83, height: 8.27 },
  'A3': { width: 11.69, height: 16.54 },
  'Custom': { width: 0, height: 0 }
};

// ========================
// MAIN FUNCTIONS
// ========================

/**
 * Creates menu when document is opened
 * This is a trigger function automatically executed by Google Docs
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Page Tools')
    .addItem('Set Custom Page Size', 'setCustomPageSize')
    .addItem('Show Current Margins', 'showCurrentMargins')
    .addToUi();
}

/**
 * Displays dialog for setting page size and formatting options
 */
function setCustomPageSize() {
  const ui = DocumentApp.getUi();
  
  // Create grouped dropdown lists for KDP sizes
  let paperbackDropdownHtml = '<optgroup label="Paperback Sizes">';
  let hardcoverDropdownHtml = '<optgroup label="Hardcover Sizes">';
  
  for (const size in KDP_SIZES) {
    const option = `<option value="${size}">${size.replace(/^(Paperback|Hardcover) - /, '')} (${KDP_SIZES[size].width}" × ${KDP_SIZES[size].height}")</option>`;
    if (KDP_SIZES[size].type === 'paperback') {
      paperbackDropdownHtml += option;
    } else {
      hardcoverDropdownHtml += option;
    }
  }
  
  paperbackDropdownHtml += '</optgroup>';
  hardcoverDropdownHtml += '</optgroup>';
  
  // Create a dropdown list HTML for common sizes
  let commonDropdownHtml = '<optgroup label="Common Sizes">';
  for (const size in COMMON_SIZES) {
    if (size !== 'Custom') {
      commonDropdownHtml += `<option value="${size}">${size} (${COMMON_SIZES[size].width}" × ${COMMON_SIZES[size].height}")</option>`;
    } else {
      commonDropdownHtml += `<option value="${size}">${size}</option>`;
    }
  }
  commonDropdownHtml += '</optgroup>';
  
  // Create the HTML for the dialog
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        select, input { width: 100%; padding: 5px; margin-bottom: 10px; }
        .button-group { text-align: right; }
        button { padding: 5px 10px; margin-left: 10px; }
        .custom-fields { display: none; }
        .note { font-size: 12px; color: #666; margin-top: 5px; }
        .info-box { background-color: #f0f0f0; padding: 10px; border-radius: 5px; margin-bottom: 15px; font-size: 12px; }
      </style>
      
      <h3>Set Page Size</h3>
      
      <div class="info-box">
        Select a KDP book size or a common paper size, or enter custom dimensions compatible with KDP requirements.
      </div>
      
      <div class="form-group">
        <label for="bookSize">Select Book Size:</label>
        <select id="bookSize">
          <option value="">-- Select a Size --</option>
          ${paperbackDropdownHtml}
          ${hardcoverDropdownHtml}
          ${commonDropdownHtml}
        </select>
      </div>
      
      <div id="customFields" class="custom-fields">
        <div class="form-group">
          <label for="customWidth">Custom Width (inches):</label>
          <input type="number" id="customWidth" step="0.01" min="${MIN_WIDTH}" max="${MAX_WIDTH}" placeholder="Width (${MIN_WIDTH}-${MAX_WIDTH} inches)">
          <div class="note">KDP paperback width must be between ${MIN_WIDTH}" and ${MAX_WIDTH}"</div>
        </div>
        
        <div class="form-group">
          <label for="customHeight">Custom Height (inches):</label>
          <input type="number" id="customHeight" step="0.01" min="${MIN_HEIGHT}" max="${MAX_HEIGHT}" placeholder="Height (${MIN_HEIGHT}-${MAX_HEIGHT} inches)">
          <div class="note">KDP paperback height must be between ${MIN_HEIGHT}" and ${MAX_HEIGHT}"</div>
        </div>
      </div>
      
      <div class="form-group">
        <label for="bookType">Book Type:</label>
        <select id="bookType">
          <option value="paperback">Paperback</option>
          <option value="hardcover">Hardcover</option>
        </select>
        <div class="note">Note: Not all sizes are available for hardcover</div>
      </div>
      
      <div class="form-group">
        <label for="paperType">Paper Type:</label>
        <select id="paperType">
          <option value="white">White</option>
          <option value="cream">Cream</option>
        </select>
      </div>
      
      <div class="form-group">
        <label for="inkType">Ink Type:</label>
        <select id="inkType">
          <option value="black">Black</option>
          <option value="premium">Premium Color</option>
          <option value="standard">Standard Color (Paperback only)</option>
        </select>
      </div>
      
      <div class="form-group">
        <label for="margins">Set Standard Margins:</label>
        <select id="margins">
          <option value="default">Default (1" all sides)</option>
          <option value="narrow">Narrow (0.5" all sides)</option>
          <option value="wide">Wide (1.25" all sides)</option>
          <option value="mirrored">Mirrored (Book-style)</option>
          <option value="custom">Custom</option>
        </select>
      </div>
      
      <div id="customMargins" style="display: none;">
        <div class="form-group">
          <label for="topMargin">Top Margin (inches):</label>
          <input type="number" id="topMargin" step="0.1" min="0.25" value="1">
        </div>
        <div class="form-group">
          <label for="bottomMargin">Bottom Margin (inches):</label>
          <input type="number" id="bottomMargin" step="0.1" min="0.25" value="1">
        </div>
        <div class="form-group">
          <label for="insideMargin">Inside/Left Margin (inches):</label>
          <input type="number" id="insideMargin" step="0.1" min="0.25" value="1">
        </div>
        <div class="form-group">
          <label for="outsideMargin">Outside/Right Margin (inches):</label>
          <input type="number" id="outsideMargin" step="0.1" min="0.25" value="1">
        </div>
      </div>
      
      <div class="button-group">
        <button id="cancelBtn" onclick="google.script.host.close()">Cancel</button>
        <button id="applyBtn" onclick="applySize()">Apply</button>
      </div>
      
      <script>
        // Show custom fields when "Custom" is selected
        document.getElementById('bookSize').addEventListener('change', function() {
          if (this.value === 'Custom') {
            document.getElementById('customFields').style.display = 'block';
          } else {
            document.getElementById('customFields').style.display = 'none';
          }
          
          // Update book type based on selection
          if (this.value.startsWith('Hardcover')) {
            document.getElementById('bookType').value = 'hardcover';
            // Disable standard color for hardcover
            updateInkOptions();
          } else if (this.value.startsWith('Paperback')) {
            document.getElementById('bookType').value = 'paperback';
            updateInkOptions();
          }
        });
        
        // Update ink options based on book type
        document.getElementById('bookType').addEventListener('change', updateInkOptions);
        
        function updateInkOptions() {
          const bookType = document.getElementById('bookType').value;
          const inkSelect = document.getElementById('inkType');
          const standardOption = inkSelect.querySelector('option[value="standard"]');
          
          if (bookType === 'hardcover') {
            standardOption.disabled = true;
            if (inkSelect.value === 'standard') {
              inkSelect.value = 'premium';
            }
          } else {
            standardOption.disabled = false;
          }
        }
        
        // Handle custom margins
        document.getElementById('margins').addEventListener('change', function() {
          if (this.value === 'custom') {
            document.getElementById('customMargins').style.display = 'block';
          } else {
            document.getElementById('customMargins').style.display = 'none';
          }
        });
        
        function applySize() {
          const selectedSize = document.getElementById('bookSize').value;
          const bookType = document.getElementById('bookType').value;
          const paperType = document.getElementById('paperType').value;
          const inkType = document.getElementById('inkType').value;
          const marginType = document.getElementById('margins').value;
          
          let width, height;
          let marginSettings = {};
          
          // Get margin settings
          if (marginType === 'custom') {
            marginSettings = {
              top: parseFloat(document.getElementById('topMargin').value),
              bottom: parseFloat(document.getElementById('bottomMargin').value),
              inside: parseFloat(document.getElementById('insideMargin').value),
              outside: parseFloat(document.getElementById('outsideMargin').value)
            };
          } else {
            marginSettings = {
              type: marginType
            };
          }
          
          if (selectedSize === 'Custom') {
            width = parseFloat(document.getElementById('customWidth').value);
            height = parseFloat(document.getElementById('customHeight').value);
            
            if (!width || !height || width < ${MIN_WIDTH} || width > ${MAX_WIDTH} || height < ${MIN_HEIGHT} || height > ${MAX_HEIGHT}) {
              alert('Please enter valid dimensions. Width must be between ${MIN_WIDTH}" and ${MAX_WIDTH}", and height between ${MIN_HEIGHT}" and ${MAX_HEIGHT}".');
              return;
            }
            
            google.script.run.withSuccessHandler(onSuccess).applyCustomPageSize(width, height, bookType, paperType, inkType, marginSettings);
          } else if (selectedSize) {
            google.script.run.withSuccessHandler(onSuccess).applyPageSettings(selectedSize, bookType, paperType, inkType, marginSettings);
          } else {
            alert('Please select a page size or enter custom dimensions.');
          }
        }
        
        function onSuccess(message) {
          alert(message);
          google.script.host.close();
        }
        
        // Initialize ink options
        updateInkOptions();
      </script>
    `)
    .setWidth(450)
    .setHeight(650);
    
  ui.showModalDialog(htmlOutput, 'Set Custom Page Size');
}

/**
 * Displays the current document margins and page size
 */
function showCurrentMargins() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Get current margins in points
  const topMargin = body.getMarginTop();
  const bottomMargin = body.getMarginBottom();
  const leftMargin = body.getMarginLeft();
  const rightMargin = body.getMarginRight();
  
  // Get current page size in points
  const pageWidth = body.getPageWidth();
  const pageHeight = body.getPageHeight();
  
  // Convert points to inches
  const topInches = topMargin / POINTS_PER_INCH;
  const bottomInches = bottomMargin / POINTS_PER_INCH;
  const leftInches = leftMargin / POINTS_PER_INCH;
  const rightInches = rightMargin / POINTS_PER_INCH;
  const widthInches = pageWidth / POINTS_PER_INCH;
  const heightInches = pageHeight / POINTS_PER_INCH;
  
  // Display the margins and page size
  const ui = DocumentApp.getUi();
  ui.alert(
    'Page Settings',
    `Page Size: ${widthInches.toFixed(2)}" × ${heightInches.toFixed(2)}" (${pageWidth} × ${pageHeight} pts)\n\n` +
    `Margins:\n` +
    `Top: ${topInches.toFixed(2)}" (${topMargin} pts)\n` +
    `Bottom: ${bottomInches.toFixed(2)}" (${bottomMargin} pts)\n` +
    `Left: ${leftInches.toFixed(2)}" (${leftMargin} pts)\n` +
    `Right: ${rightInches.toFixed(2)}" (${rightMargin} pts)`,
    ui.ButtonSet.OK
  );
}

/**
 * Applies predefined page settings based on selected size name
 * 
 * @param {string} sizeName - Name of the selected size from KDP_SIZES or COMMON_SIZES
 * @param {string} bookType - Either 'paperback' or 'hardcover'
 * @param {string} paperType - Either 'white' or 'cream'
 * @param {string} inkType - 'black', 'premium', or 'standard'
 * @param {Object} marginSettings - Margin settings object
 * @return {string} Status message
 */
function applyPageSettings(sizeName, bookType, paperType, inkType, marginSettings) {
  try {
    let widthInches, heightInches;
    
    if (KDP_SIZES[sizeName]) {
      widthInches = KDP_SIZES[sizeName].width;
      heightInches = KDP_SIZES[sizeName].height;
    } else {
      widthInches = COMMON_SIZES[sizeName].width;
      heightInches = COMMON_SIZES[sizeName].height;
    }
    
    // Convert inches to points
    const pageWidth = widthInches * POINTS_PER_INCH;
    const pageHeight = heightInches * POINTS_PER_INCH;
    
    // Apply the page size to the document
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    body.setPageWidth(pageWidth);
    body.setPageHeight(pageHeight);
    
    // Set background color based on paper type
    body.setBackgroundColor(paperType === 'cream' ? COLOR_CREAM : COLOR_WHITE);
    
    // Apply margins
    applyMargins(body, marginSettings);
    
    // Save document properties to remember settings
    saveDocumentSettings(sizeName, widthInches, heightInches, bookType, paperType, inkType);
    
    return "Page size set to " + sizeName.replace(/^(Paperback|Hardcover) - /, '') + 
           " (" + widthInches + "\" × " + heightInches + "\") as " + bookType + 
           " with " + paperType + " paper and " + inkType + " ink.";
  } catch (error) {
    Logger.log("Error in applyPageSettings: " + error);
    return "Error: " + error.toString();
  }
}

/**
 * Applies custom page size with validation
 * 
 * @param {number} widthInches - Page width in inches
 * @param {number} heightInches - Page height in inches
 * @param {string} bookType - Either 'paperback' or 'hardcover'
 * @param {string} paperType - Either 'white' or 'cream'
 * @param {string} inkType - 'black', 'premium', or 'standard'
 * @param {Object} marginSettings - Margin settings object
 * @return {string} Status message
 */
function applyCustomPageSize(widthInches, heightInches, bookType, paperType, inkType, marginSettings) {
  try {
    // Server-side validation
    if (!widthInches || !heightInches || 
        widthInches < MIN_WIDTH || widthInches > MAX_WIDTH || 
        heightInches < MIN_HEIGHT || heightInches > MAX_HEIGHT) {
      return `Error: Invalid dimensions. Width must be between ${MIN_WIDTH}" and ${MAX_WIDTH}", and height between ${MIN_HEIGHT}" and ${MAX_HEIGHT}".`;
    }
    
    // Convert inches to points
    const pageWidth = widthInches * POINTS_PER_INCH;
    const pageHeight = heightInches * POINTS_PER_INCH;
    
    // Apply the page size to the document
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    body.setPageWidth(pageWidth);
    body.setPageHeight(pageHeight);
    
    // Set background color based on paper type
    body.setBackgroundColor(paperType === 'cream' ? COLOR_CREAM : COLOR_WHITE);
    
    // Apply margins
    applyMargins(body, marginSettings);
    
    // Save document properties to remember settings
    saveDocumentSettings('Custom', widthInches, heightInches, bookType, paperType, inkType);
    
    return "Custom page size set to " + widthInches.toFixed(2) + "\" × " + heightInches.toFixed(2) + 
           "\" as " + bookType + " with " + paperType + " paper and " + inkType + " ink.";
  } catch (error) {
    Logger.log("Error in applyCustomPageSize: " + error);
    return "Error: " + error.toString();
  }
}

// ========================
// HELPER FUNCTIONS
// ========================

/**
 * Applies margin settings to document body
 * 
 * @param {Body} body - Document body object
 * @param {Object} marginSettings - Margin settings object
 */
function applyMargins(body, marginSettings) {
  if (marginSettings.type) {
    // Predefined margin settings
    switch (marginSettings.type) {
      case 'default':
        body.setMarginTop(1 * POINTS_PER_INCH);
        body.setMarginBottom(1 * POINTS_PER_INCH);
        body.setMarginLeft(1 * POINTS_PER_INCH);
        body.setMarginRight(1 * POINTS_PER_INCH);
        break;
      case 'narrow':
        body.setMarginTop(0.5 * POINTS_PER_INCH);
        body.setMarginBottom(0.5 * POINTS_PER_INCH);
        body.setMarginLeft(0.5 * POINTS_PER_INCH);
        body.setMarginRight(0.5 * POINTS_PER_INCH);
        break;
      case 'wide':
        body.setMarginTop(1.25 * POINTS_PER_INCH);
        body.setMarginBottom(1.25 * POINTS_PER_INCH);
        body.setMarginLeft(1.25 * POINTS_PER_INCH);
        body.setMarginRight(1.25 * POINTS_PER_INCH);
        break;
      case 'mirrored': // Book-style margins
        body.setMarginTop(1 * POINTS_PER_INCH);
        body.setMarginBottom(1 * POINTS_PER_INCH);
        body.setMarginLeft(1.25 * POINTS_PER_INCH); // Inside margin (for binding)
        body.setMarginRight(0.75 * POINTS_PER_INCH); // Outside margin
        break;
    }
  } else {
    // Custom margin settings
    body.setMarginTop(marginSettings.top * POINTS_PER_INCH);
    body.setMarginBottom(marginSettings.bottom * POINTS_PER_INCH);
    body.setMarginLeft(marginSettings.inside * POINTS_PER_INCH);
    body.setMarginRight(marginSettings.outside * POINTS_PER_INCH);
  }
}

/**
 * Saves current document settings to document properties
 * This allows the settings to be retrieved later
 * 
 * @param {string} sizeName - Name of the selected size
 * @param {number} width - Page width in inches
 * @param {number} height - Page height in inches
 * @param {string} bookType - Book type
 * @param {string} paperType - Paper type
 * @param {string} inkType - Ink type
 */
function saveDocumentSettings(sizeName, width, height, bookType, paperType, inkType) {
  const properties = PropertiesService.getDocumentProperties();
  const settings = {
    sizeName: sizeName,
    width: width,
    height: height,
    bookType: bookType,
    paperType: paperType,
    inkType: inkType,
    lastUpdated: new Date().toISOString()
  };
  
  properties.setProperty('kdpFormatterSettings', JSON.stringify(settings));
}
