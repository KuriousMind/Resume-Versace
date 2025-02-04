var styleCounts = {};

// ====== RULES ======
// 1. All font: Cambria
// 2. #Header1: 26pt, Centered
// 3. ##Header2: 17pt, Centered
// 4. ###Header3: 14pt, Left
// 5. "---" becomes horizontal line
// 6. Blank lines preserved
// 7. No paragraph spacing
// 8. "- " becomes "> " bullets
// 9. Bullets are indented 35pt

const STYLES = {
  header1: {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
    [DocumentApp.Attribute.FONT_SIZE]: 26,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
  },
  header2: {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
    [DocumentApp.Attribute.FONT_SIZE]: 17,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
  },
  header3: {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
    [DocumentApp.Attribute.FONT_SIZE]: 14,
    [DocumentApp.Attribute.BOLD]: true
  },
  bullet: {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
    [DocumentApp.Attribute.FONT_SIZE]: 11,
    [DocumentApp.Attribute.BOLD]: false,
    [DocumentApp.Attribute.INDENT_START]: 35,
    [DocumentApp.Attribute.INDENT_FIRST_LINE]: 35,
    [DocumentApp.Attribute.SPACING_AFTER]: 2,
  },
  horizontalRule: {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
    [DocumentApp.Attribute.PADDING_TOP]: 0,
    [DocumentApp.Attribute.PADDING_BOTTOM]: 0
  }
};

function onOpen() {
  try {
    DocumentApp.getUi()
      .createMenu('Resume Versace')
      .addItem('Ice it Out!', 'formatResumeWrapper') // Call the wrapper
      .addToUi();
    Logger.log("Menu added successfully.");
  } catch (e) {
    Logger.log("Error adding menu: " + e);
  }
}

function formatResumeWrapper() {
  let message = "Aquaberry Activated!"; // Default success message
  try {
    message = formatResume() || message; // Fallback to default if undefined
    showFormattingResult(message);
  } catch (error) {
    Logger.log('Error in formatResumeWrapper(): ' + error);
    Logger.log('Error stack: ' + error.stack);
    message = 'Hold up Versace Boss, this script done froze up more than my Iceburg Chain\n\nError details: ';
    message += error.message ? error.message : "An unknown error occurred.";
    showFormattingResult(message); // Show the error message
  }
}

function showFormattingResult(message) {
  Logger.log("showFormattingResult called with message: " + message);
  if (message) {
    DocumentApp.getUi().alert(message);
  } else {
    Logger.log("showFormattingResult called with a NULL or UNDEFINED message!");
    DocumentApp.getUi().alert("A critical error occurred. Check the logs.");
  }
}

function formatResume() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const markdown = doc.getText();
  
  body.clear();
  
  // Split on newlines but preserve empty lines
  const lines = markdown.split(/\r?\n/);
  
  let centerNextParagraph = false;  // New flag to track centering
  
  lines.forEach((originalLine) => {
    const line = originalLine.trim();
    
    try {
      // Strict header matching with regex (requires space after #)
      const headerMatch = line.match(/^(#{1,3})\s+(.*)/); // Added + for multiple spaces
      
      if (headerMatch) {
        const level = headerMatch[1].length;
        const text = headerMatch[2].trim(); // Trim header text
        
        if (level === 1) {
          const para = body.appendParagraph(text);
          para.setAttributes(STYLES.header1);
          centerNextParagraph = true;  // Set flag when H1 is found
        } 
        else if (level === 2) {
          const para = body.appendParagraph(text);
          para.setAttributes(STYLES.header2);
        }
        else if (level === 3) {
          const para = body.appendParagraph(text);
          para.setAttributes(STYLES.header3);
        }
      }
      else if (line === '---') {
        body.appendHorizontalRule();  // Add native horizontal line
      }
      else if (originalLine.startsWith('- ')) {
        const bulletText = originalLine.substring(2).trim();
        if (bulletText) {
          const bullet = body.appendParagraph('> ' + bulletText);
          bullet.setAttributes(STYLES.bullet);
        }
      }
      else {
        // Preserve original line breaks and whitespace
        const p = body.appendParagraph(originalLine);
        const attributes = {
          [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
          [DocumentApp.Attribute.FONT_SIZE]: 11,
          [DocumentApp.Attribute.BOLD]: false,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: centerNextParagraph 
            ? DocumentApp.HorizontalAlignment.CENTER  // Center if flag is set
            : DocumentApp.HorizontalAlignment.LEFT
        };
        
        p.setAttributes(attributes);
        centerNextParagraph = false;  // Reset flag after first use
      }
    } catch (e) {
      Logger.log('Error processing line: ' + originalLine + '\n' + e.stack);
    }
  });
  
  return "❄️ You ICY now ❄️";
}

function formatSection(body, section, sectionIndex) {
  const lines = (section || "").split('\n');
  let currentParagraph = null;

  // Add bold text processing
  const processBold = (paragraph, text) => {
    const boldRegex = /\*\*(.*?)\*\*/g;
    let match;
    let cleanedText = "";
    let boldRanges = [];
    let totalOffset = 0;

    // First pass: build cleaned text and track bold positions
    while ((match = boldRegex.exec(text)) !== null) {
      // Add text before the bold section
      cleanedText += text.slice(totalOffset, match.index);
      // Add bold content (without markers)
      cleanedText += match[1];
      // Calculate position in cleaned text
      const start = cleanedText.length - match[1].length;
      const end = cleanedText.length - 1;
      boldRanges.push({ start, end });
      totalOffset = match.index + match[0].length;
    }
    // Add remaining text after last match
    cleanedText += text.slice(totalOffset);

    // Set cleaned text to paragraph
    paragraph.setText(cleanedText);
    const textObj = paragraph.editAsText();

    // Apply bold formatting to tracked ranges
    boldRanges.forEach(range => {
      textObj.setBold(range.start, range.end, true);
    });
  };

  lines.forEach((line, lineIndex) => {
    let processedLine = line.trim();
    
    // Handle all header levels with same black style
    if (processedLine.startsWith('### ')) {
      const h3 = body.appendParagraph(processedLine.substring(4));
      applyStyle(h3, 'subsectionHeader');
      return;
    }
    if (processedLine.startsWith('## ')) {
      const h2 = body.appendParagraph(processedLine.substring(3));
      applyStyle(h2, 'sectionHeader');
      return;
    }
    if (processedLine.startsWith('# ')) {
      const h1 = body.appendParagraph(processedLine.substring(2));
      applyStyle(h1, 'nameHeader');
      return;
    }
    
    // Horizontal rule handling with actual line characters
    if (processedLine === "---") {
      const hr = body.appendParagraph("_________________________________________"); // Underscore-based line
      applyStyle(hr, 'decorativeHR');
      return;
    }

    // Replace bullet dashes with ">" at start of lines
    if (processedLine.startsWith("- ")) {
      processedLine = "> " + processedLine.substring(2);
    }

    try {
      if (processedLine.startsWith('## ')) { // H2
        currentParagraph = body.appendParagraph(processedLine.substring(3));
        applyStyle(currentParagraph, 'subsectionHeader');
      } else if (processedLine.startsWith('### ')) { // H3
        currentParagraph = body.appendParagraph(processedLine.substring(4));
        applyStyle(currentParagraph, 'subCompanyHeader');
      } else if (processedLine !== "") { // Regular text
        currentParagraph = body.appendParagraph('');
        processBold(currentParagraph, processedLine);
        
        // Apply different style for bullet points
        if (processedLine.startsWith("> ")) {
          applyStyle(currentParagraph, 'bulletPoint');
        } else {
          applyStyle(currentParagraph, 'regularText');
        }
      } else if (processedLine === "") { // Handle empty lines
        // Add spacer paragraph for empty lines
        const spacer = body.appendParagraph("");
        spacer.setSpacingBefore(12); // Space before next paragraph
        spacer.setSpacingAfter(0);
        currentParagraph = null;
      }
    } catch (lineError) {
      // ... (error handling)
    }
  });
}

/**
 * Applies a predefined style to a paragraph.
 * @param {DocumentApp.Paragraph} paragraph The paragraph to style.
 * @param {string} styleName The name of the style to apply (e.g., 'heading1').
 */
function applyStyle(paragraph, styleName) {
  if (!styleName || typeof styleName !== 'string') {
    Logger.log("Invalid style name: " + styleName);
    return;
  }

  const styles = {
    nameHeader: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
      [DocumentApp.Attribute.FONT_SIZE]: 26,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000' // Black color
    },
    sectionHeader: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
      [DocumentApp.Attribute.FONT_SIZE]: 17,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000'
    },
    subsectionHeader: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
      [DocumentApp.Attribute.FONT_SIZE]: 14,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15
    },
    companyHeader: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
      [DocumentApp.Attribute.FONT_SIZE]: 13,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15,
      [DocumentApp.Attribute.SPACING_BEFORE]: 10
    },
    subCompanyHeader: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri',
      [DocumentApp.Attribute.FONT_SIZE]: 14,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.ITALIC]: true,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#00b894',
      [DocumentApp.Attribute.SPACING_BEFORE]: 12
    },
    bulletPoint: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri',
      [DocumentApp.Attribute.FONT_SIZE]: 11,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000',
      [DocumentApp.Attribute.INDENT_START]: 36,
      [DocumentApp.Attribute.INDENT_FIRST_LINE]: -18,
      [DocumentApp.Attribute.SPACING_AFTER]: 6,
      [DocumentApp.Attribute.LINE_SPACING]: 1,
      [DocumentApp.Attribute.BULLET]: true
    },
    expertiseHeader: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
      [DocumentApp.Attribute.FONT_SIZE]: 11,
      [DocumentApp.Attribute.BOLD]: false,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15,
      [DocumentApp.Attribute.SPACING_BEFORE]: 10
    },
    name: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
      [DocumentApp.Attribute.FONT_SIZE]: 28,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15
    },
    regularText: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
      [DocumentApp.Attribute.FONT_SIZE]: 11,
      [DocumentApp.Attribute.BOLD]: false,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15
    },
    contactLine: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri',
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#636e72',
      [DocumentApp.Attribute.SPACING_AFTER]: 24
    },
    decorativeHR: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri',
      [DocumentApp.Attribute.FONT_SIZE]: 14,
      [DocumentApp.Attribute.BOLD]: false,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#2d3436', // Darker gray
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
      [DocumentApp.Attribute.LINE_SPACING]: 1.3,
    }
  };

  Logger.log("Styles object: " + JSON.stringify(styles, null, 2));

  if (!paragraph) {  // Check if paragraph is initially null
    Logger.log("Paragraph is initially null. Cannot apply style: " + styleName);
    return; // Early exit
  }

  if (styles[styleName]) {
    try {
      Logger.log("Applying style: " + styleName + " to paragraph: " + paragraph.getText());
      
      // Create copy of style attributes
      const styleAttributes = Object.assign({}, styles[styleName]);
      
      // Modified centering logic
      if (['sectionHeader'].includes(styleName)) {  // ONLY SECTION HEADER GETS CENTERING
        styleCounts[styleName] = (styleCounts[styleName] || 0) + 1;
        if (styleCounts[styleName] === 1) {
          styleAttributes[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = 
            DocumentApp.HorizontalAlignment.CENTER;
        } else {
          styleAttributes[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = 
            DocumentApp.HorizontalAlignment.LEFT;
        }
      }

      paragraph.setAttributes(styleAttributes);
      Logger.log("Paragraph text after styling: " + paragraph.getText());

      // Check if paragraph became null AFTER setAttributes (unlikely, but good to check)
      if (!paragraph) {
        Logger.log("Paragraph became null AFTER setAttributes. This indicates a serious problem.");
        return; // Exit to prevent further errors
      }
    } catch (e) {
      Logger.log(`Error applying style ${styleName} to paragraph: ${e}`);
      Logger.log(`Error stack: ${e.stack}`);
      // Consider what action to take here.  You might want to return,
      // or you might want to try to continue processing other paragraphs.
      return; // Or remove this if you want the script to continue
    }
  } else {
    Logger.log("Style '" + styleName + "' not found in styles object.");
  }
}