// ============= Configuration =============
const IMRAD_STRUCTURE = {
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

// ============= UI Helper =============
class UIHelper {
  // ... UI related methods
}

// ============= Document Manager =============
class DocumentManager {
  // ... document manipulation methods
}

// ============= Main Application =============
function onOpen() {
  DocumentApp.getUi()
    .createMenu('IMRAD Helper')
    .addItem('Create IMRAD Structure', 'showSectionSelector')
    .addItem('Analyze Document Structure', 'analyzeDocument')
    .addToUi();
}

function showSectionSelector() {
  // ... show selector dialog
}

function insertSelectedStructure(selections) {
  // ... handle structure insertion
}

function analyzeDocument() {
  // ... handle document analysis
}

// IMRAD Structure Helper
class IMRADHelper {
  constructor() {
    this.sections = IMRAD_STRUCTURE;
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

// Add this function to show the selection dialog
function showSectionSelector() {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
          }
          .section {
            margin-bottom: 20px;
          }
          .section-title {
            font-weight: bold;
            margin-bottom: 10px;
          }
          .subsection {
            margin-left: 20px;
            margin-bottom: 5px;
          }
          .button-container {
            margin-top: 20px;
            text-align: right;
          }
          button {
            padding: 8px 16px;
            margin-left: 10px;
          }
          .select-all {
            margin-bottom: 15px;
          }
        </style>
      </head>
      <body>
        <h3>Select IMRAD Sections</h3>
        <div class="select-all">
          <input type="checkbox" id="selectAll" checked>
          <label for="selectAll">Select/Deselect All</label>
        </div>
        
        <div class="sections">
          <div class="section">
            <div class="section-title">
              <input type="checkbox" id="introduction" checked>
              <label for="introduction">Introduction</label>
            </div>
            <div class="subsections" id="introduction-subs">
              <div class="subsection">
                <input type="checkbox" id="background" checked>
                <label for="background">Background</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="problem-statement" checked>
                <label for="problem-statement">Problem Statement</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="research-objectives" checked>
                <label for="research-objectives">Research Objectives</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="research-questions" checked>
                <label for="research-questions">Research Questions/Hypotheses</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="significance" checked>
                <label for="significance">Significance of the Study</label>
              </div>
            </div>
          </div>

          <div class="section">
            <div class="section-title">
              <input type="checkbox" id="methods" checked>
              <label for="methods">Methods</label>
            </div>
            <div class="subsections" id="methods-subs">
              <div class="subsection">
                <input type="checkbox" id="research-design" checked>
                <label for="research-design">Research Design</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="participants" checked>
                <label for="participants">Participants/Sample</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="data-collection" checked>
                <label for="data-collection">Data Collection</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="data-analysis" checked>
                <label for="data-analysis">Data Analysis</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="ethical" checked>
                <label for="ethical">Ethical Considerations</label>
              </div>
            </div>
          </div>

          <div class="section">
            <div class="section-title">
              <input type="checkbox" id="results" checked>
              <label for="results">Results</label>
            </div>
            <div class="subsections" id="results-subs">
              <div class="subsection">
                <input type="checkbox" id="data-presentation" checked>
                <label for="data-presentation">Data Presentation</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="statistical-analysis" checked>
                <label for="statistical-analysis">Statistical Analysis</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="key-findings" checked>
                <label for="key-findings">Key Findings</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="tables-figures" checked>
                <label for="tables-figures">Tables and Figures</label>
              </div>
            </div>
          </div>

          <div class="section">
            <div class="section-title">
              <input type="checkbox" id="discussion" checked>
              <label for="discussion">Discussion</label>
            </div>
            <div class="subsections" id="discussion-subs">
              <div class="subsection">
                <input type="checkbox" id="interpretation" checked>
                <label for="interpretation">Interpretation of Results</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="implications" checked>
                <label for="implications">Implications</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="limitations" checked>
                <label for="limitations">Limitations</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="future-research" checked>
                <label for="future-research">Future Research</label>
              </div>
              <div class="subsection">
                <input type="checkbox" id="conclusion" checked>
                <label for="conclusion">Conclusion</label>
              </div>
            </div>
          </div>
        </div>

        <div class="button-container">
          <button onclick="google.script.host.close()">Cancel</button>
          <button onclick="submitSelections()">Create Structure</button>
        </div>

        <script>
          // Handle select all checkbox
          document.getElementById('selectAll').addEventListener('change', function(e) {
            const checkboxes = document.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(checkbox => {
              checkbox.checked = e.target.checked;
            });
          });

          // Handle section checkbox to toggle its subsections
          document.querySelectorAll('.section-title input').forEach(checkbox => {
            checkbox.addEventListener('change', function(e) {
              const subsections = this.closest('.section').querySelectorAll('.subsections input');
              subsections.forEach(sub => {
                sub.checked = e.target.checked;
              });
            });
          });

          function submitSelections() {
            const selections = {
              introduction: {
                selected: document.getElementById('introduction').checked,
                subsections: {
                  background: document.getElementById('background').checked,
                  problemStatement: document.getElementById('problem-statement').checked,
                  researchObjectives: document.getElementById('research-objectives').checked,
                  researchQuestions: document.getElementById('research-questions').checked,
                  significance: document.getElementById('significance').checked
                }
              },
              methods: {
                selected: document.getElementById('methods').checked,
                subsections: {
                  researchDesign: document.getElementById('research-design').checked,
                  participants: document.getElementById('participants').checked,
                  dataCollection: document.getElementById('data-collection').checked,
                  dataAnalysis: document.getElementById('data-analysis').checked,
                  ethical: document.getElementById('ethical').checked
                }
              },
              results: {
                selected: document.getElementById('results').checked,
                subsections: {
                  dataPresentation: document.getElementById('data-presentation').checked,
                  statisticalAnalysis: document.getElementById('statistical-analysis').checked,
                  keyFindings: document.getElementById('key-findings').checked,
                  tablesFigures: document.getElementById('tables-figures').checked
                }
              },
              discussion: {
                selected: document.getElementById('discussion').checked,
                subsections: {
                  interpretation: document.getElementById('interpretation').checked,
                  implications: document.getElementById('implications').checked,
                  limitations: document.getElementById('limitations').checked,
                  futureResearch: document.getElementById('future-research').checked,
                  conclusion: document.getElementById('conclusion').checked
                }
              }
            };
            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .insertSelectedStructure(selections);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(600)
    .setTitle('Select IMRAD Sections');

  DocumentApp.getUi().showModalDialog(htmlOutput, 'Select IMRAD Sections');
}

// New function to handle the selected sections
function insertSelectedStructure(selections) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Clear existing content
  body.clear();
  
  // Set default paragraph style
  const style = {};
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman';
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  body.setAttributes(style);
  
  // Insert title
  const title = body.appendParagraph('Title');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  title.setAttributes(style);

  // Insert selected sections
  if (selections.introduction.selected) {
    insertSection('Introduction', selections.introduction.subsections, body, style);
  }
  if (selections.methods.selected) {
    insertSection('Methods', selections.methods.subsections, body, style);
  }
  if (selections.results.selected) {
    insertSection('Results', selections.results.subsections, body, style);
  }
  if (selections.discussion.selected) {
    insertSection('Discussion', selections.discussion.subsections, body, style);
  }
}

// Helper function to insert a section and its selected subsections
function insertSection(sectionTitle, subsections, body, style) {
  const sectionPara = body.appendParagraph(sectionTitle);
  sectionPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  sectionPara.setAttributes(style);

  // Get the section's subsections and placeholders from the original structure
  const sectionData = IMRAD_STRUCTURE[sectionTitle.toLowerCase()];
  
  sectionData.subsections.forEach(subsection => {
    const subsectionId = subsection.toLowerCase().replace(/\s+/g, '-');
    if (subsections[subsectionId]) {
      const subsectionPara = body.appendParagraph(subsection);
      subsectionPara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      subsectionPara.setAttributes(style);
      
      const placeholderPara = body.appendParagraph(subsection.placeholder);
      placeholderPara.setAttributes(style);
    }
  });

  // Add spacing between sections
  body.appendParagraph('').setAttributes(style);
}