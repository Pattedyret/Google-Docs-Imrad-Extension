// IMRAD Structure Helper
class IMRADHelper {
  constructor() {
    this.sections = {
      introduction: {
        title: 'Introduction',
        subsections: [
          'Background',
          'Problem Statement',
          'Research Objectives',
          'Research Questions/Hypotheses',
          'Significance of the Study'
        ]
      },
      methods: {
        title: 'Methods',
        subsections: [
          'Research Design',
          'Participants/Sample',
          'Data Collection',
          'Data Analysis',
          'Ethical Considerations'
        ]
      },
      results: {
        title: 'Results',
        subsections: [
          'Data Presentation',
          'Statistical Analysis',
          'Key Findings',
          'Tables and Figures'
        ]
      },
      discussion: {
        title: 'Discussion',
        subsections: [
          'Interpretation of Results',
          'Implications',
          'Limitations',
          'Future Research',
          'Conclusion'
        ]
      }
    };
  }

  insertIMRADStructure() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    // Clear existing content
    body.clear();
    
    // Insert title
    body.appendParagraph('Title').setHeading(DocumentApp.ParagraphHeading.TITLE);
    
    // Insert each section
    for (const [key, section] of Object.entries(this.sections)) {
      // Insert main section heading
      const sectionPara = body.appendParagraph(section.title);
      sectionPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      
      // Insert subsections
      section.subsections.forEach(subsection => {
        const subsectionPara = body.appendParagraph(subsection);
        subsectionPara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        // Add placeholder text
        body.appendParagraph('Add your content here...');
      });
      
      // Add spacing between sections
      body.appendParagraph('');
    }

    // Show confirmation
    DocumentApp.getUi().alert('IMRAD structure has been inserted!');
  }

  analyzeDocument() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const paragraphs = body.getParagraphs();
    
    // Create a checklist of found components
    const checklist = {
      introduction: {
        title: 'Introduction',
        found: false,
        subsections: {}
      },
      methods: {
        title: 'Methods',
        found: false,
        subsections: {}
      },
      results: {
        title: 'Results',
        found: false,
        subsections: {}
      },
      discussion: {
        title: 'Discussion',
        found: false,
        subsections: {}
      }
    };

    // Initialize subsections
    for (const [key, section] of Object.entries(this.sections)) {
      section.subsections.forEach(subsection => {
        checklist[key].subsections[subsection] = false;
      });
    }

    // Analyze each paragraph
    paragraphs.forEach(paragraph => {
      const text = paragraph.getText().trim();
      
      // Check for main sections
      for (const [key, section] of Object.entries(this.sections)) {
        if (text.toLowerCase().includes(section.title.toLowerCase())) {
          checklist[key].found = true;
        }
        
        // Check for subsections
        section.subsections.forEach(subsection => {
          if (text.toLowerCase().includes(subsection.toLowerCase())) {
            checklist[key].subsections[subsection] = true;
          }
        });
      }
    });

    this.showAnalysisReport(checklist);
  }

  showAnalysisReport(checklist) {
    let report = 'IMRAD Structure Analysis\n\n';
    
    for (const [key, section] of Object.entries(checklist)) {
      const sectionStatus = section.found ? '✓' : '✗';
      report += `${section.title} ${sectionStatus}\n`;
      
      for (const [subsection, found] of Object.entries(section.subsections)) {
        const subsectionStatus = found ? '✓' : '✗';
        report += `  ${subsection} ${subsectionStatus}\n`;
      }
      report += '\n';
    }

    DocumentApp.getUi().alert(report);
  }
}

// Initialize the helper
const imradHelper = new IMRADHelper();

// Add menu item to Google Docs
function onOpen() {
  DocumentApp.getUi()
    .createMenu('IMRAD Helper')
    .addItem('Insert IMRAD Structure', 'insertIMRADStructure')
    .addItem('Analyze Document Structure', 'analyzeDocument')
    .addToUi();
}

// Function to be called from menu
function insertIMRADStructure() {
  imradHelper.insertIMRADStructure();
}

function analyzeDocument() {
  imradHelper.analyzeDocument();
}