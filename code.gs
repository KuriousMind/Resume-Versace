/**
 * Adds custom menu when document is opened - uses simple trigger
 */
function onOpen() {
  const ui = PropertiesService.getDocumentProperties();
  try {
    DocumentApp.getUi()
      .createMenu('Resume Versace')
      .addItem('Ice it Out!', 'formatResume')
      .addItem('Test H1', 'testH1WithHeading')
      .addToUi();
  } catch(e) {
    // Silent fail for auth trigger
  }
}

/**
 * Test function for H1 formatting
 */
function testH1WithHeading() {
  try {
    Logger.log('Starting test function');
    
    const doc = DocumentApp.getActiveDocument();
    Logger.log('Got document');
    
    const body = doc.getBody();
    Logger.log('Got body');
    
    const paragraphs = body.getParagraphs();
    Logger.log('Got paragraphs: ' + paragraphs.length);
    
    // Log first paragraph content
    if (paragraphs.length > 0) {
      const firstPara = paragraphs[0];
      const text = firstPara.getText().trim();
      Logger.log('First paragraph text: "' + text + '"');
      
      // If it's our name with MBA
      if (text.includes('MBA')) {
        Logger.log('Found our H1!');
        Logger.log('Current font size: ' + firstPara.getFontSize());
        
        // Try direct method
        firstPara.setFontSize(26);
        Logger.log('Set font size to 26');
        Logger.log('New font size: ' + firstPara.getFontSize());
        
        // Also try setting as attribute
        firstPara.setAttributes({
          [DocumentApp.Attribute.FONT_SIZE]: 26,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria'
        });
        Logger.log('Set via attributes');
        Logger.log('Final font size: ' + firstPara.getFontSize());
      } else {
        Logger.log('First paragraph is not H1');
      }
    } else {
      Logger.log('No paragraphs found');
    }
    
    Logger.log('Test completed successfully');
    DocumentApp.getUi().alert('Test completed - check execution log');
    
  } catch (error) {
    Logger.log('Error occurred: ' + error.message);
    DocumentApp.getUi().alert('Error: ' + error.message);
  }
}

/**
 * Main function to format the resume
 */
function formatResume() {
  try {
    const doc = DocumentApp.getActiveDocument();
    if (!doc) {
      throw new Error('No active document found');
    }

    const body = doc.getBody();
    if (!body) {
      throw new Error('Could not access document body');
    }
    
    // Define styles object
    const styles = {
      introHeader: {
        [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
        [DocumentApp.Attribute.FONT_SIZE]: 14,
        [DocumentApp.Attribute.BOLD]: true,
        [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
        [DocumentApp.Attribute.LINE_SPACING]: 1.15
      },
      introText: {
        [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
        [DocumentApp.Attribute.FONT_SIZE]: 11,
        [DocumentApp.Attribute.BOLD]: false,
        [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
        [DocumentApp.Attribute.LINE_SPACING]: 1.15
      },
      expertiseSection: {
        [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
        [DocumentApp.Attribute.FONT_SIZE]: 11,
        [DocumentApp.Attribute.BOLD]: false,
        [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
        [DocumentApp.Attribute.LINE_SPACING]: 1.15
      },
      mainHeader: {
        [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
        [DocumentApp.Attribute.FONT_SIZE]: 26,  // Set to 26 as requested
        [DocumentApp.Attribute.BOLD]: true,
        [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
        [DocumentApp.Attribute.LINE_SPACING]: 1.15
      },
      companyHeader: {
        [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
        [DocumentApp.Attribute.FONT_SIZE]: 13,
        [DocumentApp.Attribute.BOLD]: true,
        [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
        [DocumentApp.Attribute.LINE_SPACING]: 1.15
      },
      bulletPoint: {
        [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
        [DocumentApp.Attribute.FONT_SIZE]: 11,
        [DocumentApp.Attribute.BOLD]: false,
        [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
        [DocumentApp.Attribute.LINE_SPACING]: 1.15
      }
    };
    
    // Get all paragraphs
    const paragraphs = body.getParagraphs();
    
    Logger.log('Total paragraphs found: ' + paragraphs.length);
    
    // Process each paragraph
    let foundH1 = false;
    
    paragraphs.forEach((para, index) => {
      let text = para.getText().trim();
      
      // Clear existing formatting
      para.setAttributes({});
      
      // Heading 1 (starts with single #)
      if (text.startsWith('# ')) {
        Logger.log('Found H1: ' + text);
        text = text.replace(/^# /, '');
        para.setText(text);
        // Apply heading 1 style first
        para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
        // Then apply our specific formatting
        para.setAttributes({
          [DocumentApp.Attribute.FONT_FAMILY]: 'Cambria',
          [DocumentApp.Attribute.FONT_SIZE]: 26,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.LINE_SPACING]: 1.15
        });
        Logger.log('Applied H1 style');
      }
      // Major section headers
      else if (text === 'PROFESSIONAL EXPERIENCE' || 
          text === 'EDUCATION' || 
          text === 'CERTIFICATIONS & TRAINING' || 
          text === 'COMPETENCIES & TECHNICAL SKILLS') {
        para.setAttributes(styles.mainHeader);
        
        // Add spacing before major sections
        if (index > 0) {
          para.setSpacingBefore(20);
        }
      }
      // Introduction section (starts with ##)
      else if (text.startsWith('## ')) {
        text = text.replace(/^## /, '');
        para.setText(text);
        para.setAttributes(styles.introHeader);
      }
      // Company headers (starts with ###)
      else if (text.startsWith('### ')) {
        text = text.replace(/^### /, '');
        para.setText(text);
        para.setAttributes(styles.companyHeader);
        
        // Add spacing before company sections
        para.setSpacingBefore(10);
      }
      // Areas of expertise section
      else if (text === 'Areas of Expertise:') {
        para.setAttributes(styles.expertiseSection);
        para.setSpacingBefore(10);
      }
      // Bullet points
      else if (text.startsWith('-')) {
        // Ensure proper dash formatting without converting to bullet
        if (!text.startsWith('- ')) {
          text = '- ' + text.substring(1).trim();
          para.setText(text);
        }
        para.setAttributes(styles.bulletPoint);
        para.setIndentStart(36);
      }
      // Regular text (intro paragraph, etc.)
      else {
        para.setAttributes(styles.introText);
      }
      
      // Special handling for expertise bullet points (inline)
      if (text.includes('*') && text.split('*').length > 2) {
        // Format the line with proper bullet spacing
        text = text.replace(/\s*\*\s*/g, ' • ');
        para.setText(text);
      }
    });
    
    // Success message
    DocumentApp.getUi().alert('Aquaberry Activated!');
    
  } catch (error) {
    // Show error message
    DocumentApp.getUi().alert('Hold up Versace Boss, this script done froze up more than my Iceburg Chain\n\nError details: ' + error.message);
    Logger.log('Error: ' + error.toString());
  }
}