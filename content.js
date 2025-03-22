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

  async insertIMRADStructure() {
    const doc = await this.getCurrentDocument();
    const body = doc.body;
    
    // Clear existing content
    body.clear();
    
    // Insert title
    body.appendParagraph('Title').setHeading(DocumentApp.ParagraphHeading.TITLE);
    
    // Insert each section
    for (const [key, section] of Object.entries(this.sections)) {
      // Insert main section heading
      const sectionPara = body.appendParagraph(section.title);
      sectionPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      sectionPara.setId(key); // Add ID for navigation
      
      // Insert subsections
      section.subsections.forEach(subsection => {
        const subsectionPara = body.appendParagraph(subsection);
        subsectionPara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        subsectionPara.setId(`${key}-${subsection.toLowerCase().replace(/\s+/g, '-')}`); // Add ID for navigation
        // Add placeholder text
        body.appendParagraph('Add your content here...');
      });
      
      // Add spacing between sections
      body.appendParagraph('');
    }
  }

  async analyzeDocument() {
    const doc = await this.getCurrentDocument();
    const body = doc.body;
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
      const heading = paragraph.getHeading();
      const id = paragraph.getId();
      
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

    // Create analysis report
    const report = this.generateAnalysisReport(checklist);
    this.showAnalysisReport(report);
  }

  generateAnalysisReport(checklist) {
    let report = {
      sections: [],
      completion: 0
    };
    
    for (const [key, section] of Object.entries(checklist)) {
      const sectionData = {
        title: section.title,
        found: section.found,
        id: key,
        subsections: []
      };
      
      for (const [subsection, found] of Object.entries(section.subsections)) {
        sectionData.subsections.push({
          title: subsection,
          found: found,
          id: `${key}-${subsection.toLowerCase().replace(/\s+/g, '-')}`
        });
      }
      
      report.sections.push(sectionData);
    }

    // Calculate completion percentage
    const totalComponents = Object.values(checklist).reduce((acc, section) => {
      return acc + 1 + Object.keys(section.subsections).length;
    }, 0);
    
    const completedComponents = Object.values(checklist).reduce((acc, section) => {
      return acc + (section.found ? 1 : 0) + 
        Object.values(section.subsections).filter(found => found).length;
    }, 0);
    
    report.completion = Math.round((completedComponents / totalComponents) * 100);
    
    return report;
  }

  showAnalysisReport(report) {
    const html = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              color: #333;
            }
            .section {
              margin-bottom: 20px;
              padding: 10px;
              border-radius: 5px;
              background-color: #f5f5f5;
            }
            .section-title {
              font-weight: bold;
              margin-bottom: 10px;
              color: #2c3e50;
              cursor: pointer;
            }
            .section-title:hover {
              text-decoration: underline;
            }
            .subsection {
              margin-left: 20px;
              margin-bottom: 5px;
              cursor: pointer;
            }
            .subsection:hover {
              text-decoration: underline;
            }
            .found {
              color: #27ae60;
            }
            .missing {
              color: #e74c3c;
            }
            .completion {
              text-align: center;
              font-size: 1.2em;
              margin: 20px 0;
              padding: 10px;
              background-color: #ecf0f1;
              border-radius: 5px;
            }
            .progress-bar {
              width: 100%;
              height: 20px;
              background-color: #bdc3c7;
              border-radius: 10px;
              overflow: hidden;
              margin-top: 10px;
            }
            .progress {
              height: 100%;
              background-color: #3498db;
              transition: width 0.3s ease;
            }
            .back-to-top {
              position: fixed;
              bottom: 20px;
              right: 20px;
              background-color: #3498db;
              color: white;
              padding: 10px 20px;
              border-radius: 5px;
              cursor: pointer;
              box-shadow: 0 2px 5px rgba(0,0,0,0.2);
              transition: background-color 0.3s ease;
            }
            .back-to-top:hover {
              background-color: #2980b9;
            }
            .customize-button {
              background-color: #2ecc71;
              color: white;
              padding: 10px 20px;
              border-radius: 5px;
              cursor: pointer;
              margin-bottom: 20px;
              width: 100%;
              text-align: center;
              transition: background-color 0.3s ease;
            }
            .customize-button:hover {
              background-color: #27ae60;
            }
          </style>
          <script>
            function scrollToElement(id) {
              google.script.run
                .withSuccessHandler(() => {
                  // Highlight the element briefly
                  const element = document.getElementById(id);
                  if (element) {
                    element.style.backgroundColor = '#fff3cd';
                    setTimeout(() => {
                      element.style.backgroundColor = '';
                    }, 2000);
                  }
                })
                .scrollToElement(id);
            }

            function scrollToTop() {
              google.script.run.scrollToElement('title');
            }

            function showCustomizeDialog() {
              google.script.run.showCustomizeDialog();
            }
          </script>
        </head>
        <body>
          <div class="customize-button" onclick="showCustomizeDialog()">
            Customize IMRAD Structure
          </div>
          
          <div class="completion">
            <div>Overall Completion</div>
            <div class="progress-bar">
              <div class="progress" style="width: ${report.completion}%"></div>
            </div>
            <div>${report.completion}%</div>
          </div>
          
          ${report.sections.map(section => `
            <div class="section">
              <div class="section-title" onclick="scrollToElement('${section.id}')">
                ${section.title} ${section.found ? '✅' : '❌'}
              </div>
              ${section.subsections.map(subsection => `
                <div class="subsection" onclick="scrollToElement('${subsection.id}')">
                  ${subsection.title} ${subsection.found ? '✅' : '❌'}
                </div>
              `).join('')}
            </div>
          `).join('')}

          <div class="back-to-top" onclick="scrollToTop()">
            Back to Top
          </div>
        </body>
      </html>
    `)
    .setTitle('IMRAD Structure Analysis')
    .setWidth(300);

    DocumentApp.getUi().showSidebar(html);
  }

  async getCurrentDocument() {
    return new Promise((resolve) => {
      google.script.run
        .withSuccessHandler(resolve)
        .getActiveDocument();
    });
  }
}

// Initialize the helper
const imradHelper = new IMRADHelper();

// Add menu item to Google Docs
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('IMRAD Helper')
    .addItem('Insert IMRAD Structure', 'insertIMRADStructure')
    .addItem('Analyze Document Structure', 'analyzeDocument')
    .addItem('Customize Structure', 'showCustomizeDialog')
    .addToUi();
}

// Function to be called from menu
function insertIMRADStructure() {
  imradHelper.insertIMRADStructure();
}

function analyzeDocument() {
  imradHelper.analyzeDocument();
}

// Function to scroll to an element
function scrollToElement(elementId) {
  const doc = DocumentApp.getActiveDocument();
  const element = doc.getElementById(elementId);
  if (element) {
    element.scrollIntoView();
  }
}

// Function to show customize dialog
function showCustomizeDialog() {
  const ui = DocumentApp.getUi();
  const sections = imradHelper.sections;
  
  let html = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>Customize IMRAD Structure</h2>
      <p>Select which sections and subsections you want to include:</p>
      <form id="customizeForm">
  `;

  // Add checkboxes for each section and subsection
  for (const [key, section] of Object.entries(sections)) {
    html += `
      <div style="margin-bottom: 15px;">
        <input type="checkbox" id="${key}" name="${key}" checked>
        <label for="${key}"><strong>${section.title}</strong></label>
        <div style="margin-left: 20px;">
    `;
    
    section.subsections.forEach(subsection => {
      html += `
        <div>
          <input type="checkbox" id="${key}-${subsection}" name="${key}-${subsection}" checked>
          <label for="${key}-${subsection}">${subsection}</label>
        </div>
      `;
    });
    
    html += `
        </div>
      </div>
    `;
  }

  html += `
        <div style="margin-top: 20px;">
          <button type="button" onclick="saveCustomization()">Save Changes</button>
          <button type="button" onclick="google.script.host.close()">Cancel</button>
        </div>
      </form>
    </div>
    <script>
      function saveCustomization() {
        const form = document.getElementById('customizeForm');
        const formData = new FormData(form);
        const customization = {};
        
        for (const [key, value] of formData.entries()) {
          customization[key] = value === 'on';
        }
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .saveCustomization(customization);
      }
    </script>
  `;

  const dialog = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(500)
    .setTitle('Customize IMRAD Structure');

  ui.showModalDialog(dialog, 'Customize IMRAD Structure');
}

// Function to save customization
function saveCustomization(customization) {
  const doc = DocumentApp.getActiveDocument();
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('imradCustomization', JSON.stringify(customization));
  
  // Update the sections based on customization
  for (const [key, section] of Object.entries(imradHelper.sections)) {
    if (!customization[key]) {
      delete imradHelper.sections[key];
    } else {
      // Filter subsections based on customization
      section.subsections = section.subsections.filter(subsection => 
        customization[`${key}-${subsection}`]
      );
    }
  }
  
  // Show confirmation
  const ui = DocumentApp.getUi();
  ui.alert('IMRAD structure has been customized successfully!');
} 