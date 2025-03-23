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
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
            padding: 24px;
            background-color: #f8f9fa;
            color: #202124;
            max-width: 800px;
            margin: 0 auto;
          }

          .card {
            background: white;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.24);
            padding: 20px;
            margin-bottom: 24px;
            transition: all 0.3s cubic-bezier(.25,.8,.25,1);
          }

          .card:hover {
            box-shadow: 0 3px 6px rgba(0,0,0,0.16), 0 3px 6px rgba(0,0,0,0.23);
          }

          h3 {
            color: #1a73e8;
            margin-top: 0;
            font-size: 18px;
            font-weight: 500;
            margin-bottom: 16px;
            border-bottom: 2px solid #e8f0fe;
            padding-bottom: 8px;
          }

          select, input[type="text"] {
            width: 100%;
            padding: 8px 12px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-size: 14px;
            margin: 4px 0;
            transition: border-color 0.2s;
          }

          select:focus, input[type="text"]:focus {
            outline: none;
            border-color: #1a73e8;
          }

          .checkbox-wrapper {
            display: flex;
            align-items: center;
            margin: 8px 0;
            padding: 4px 0;
          }

          input[type="checkbox"] {
            margin-right: 8px;
            width: 18px;
            height: 18px;
            accent-color: #1a73e8;
          }

          label {
            font-size: 14px;
            color: #5f6368;
          }

          .section {
            background: #f8f9fa;
            border-radius: 4px;
            padding: 16px;
            margin: 12px 0;
            border: 1px solid #dadce0;
          }

          .subsection {
            margin-left: 24px;
            padding: 8px 0;
            display: flex;
            align-items: center;
          }

          .subsection input[type="text"] {
            flex-grow: 1;
            margin-left: 8px;
          }

          button {
            background: #1a73e8;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            font-size: 14px;
            cursor: pointer;
            transition: background 0.2s;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 100px;
          }

          button:hover {
            background: #1557b0;
          }

          .secondary-button {
            background: #fff;
            color: #1a73e8;
            border: 1px solid #1a73e8;
            margin-right: 8px;
          }

          .secondary-button:hover {
            background: #e8f0fe;
          }

          .add-subsection-btn {
            background: #34a853;
            margin-top: 8px;
            margin-left: 24px;
          }

          .add-subsection-btn:hover {
            background: #2d8544;
          }

          .preset-info {
            font-size: 13px;
            color: #5f6368;
            margin-top: 8px;
            padding: 8px;
            background: #e8f0fe;
            border-radius: 4px;
            border-left: 4px solid #1a73e8;
          }

          .button-container {
            display: flex;
            justify-content: flex-end;
            margin-top: 24px;
            padding-top: 16px;
            border-top: 1px solid #dadce0;
          }

          .template-actions {
            display: flex;
            gap: 8px;
            margin-top: 12px;
          }

          .template-actions select {
            flex-grow: 1;
          }

          .save-template {
            background: #34a853;
          }

          .save-template:hover {
            background: #2d8544;
          }

          .load-template {
            background: #ea4335;
          }

          .load-template:hover {
            background: #d33828;
          }
        </style>
      </head>
      <body>
        <div class="card">
          <h3>Select Report Type</h3>
          <select id="presetType" onchange="loadPreset()">
            <option value="standard">Standard Research Paper</option>
            <option value="systematic">Systematic Review</option>
            <option value="case">Case Study</option>
            <option value="experimental">Experimental Research</option>
            <option value="qualitative">Qualitative Research</option>
            <option value="mixed">Mixed Methods Research</option>
            <option value="custom">Custom Template</option>
          </select>
          <div id="presetInfo" class="preset-info"></div>
        </div>

        <div class="card">
          <h3>Templates</h3>
          <div class="template-actions">
            <select id="savedTemplates">
              <option value="">Select a saved template...</option>
            </select>
            <button onclick="loadSavedTemplate()" class="load-template">Load</button>
            <button onclick="showSaveDialog()" class="save-template">Save</button>
          </div>
        </div>

        <div class="card">
          <h3>Document Settings</h3>
          <div>
            <label for="fontFamily">Font:</label>
            <select id="fontFamily">
              <option value="Times New Roman">Times New Roman</option>
              <option value="Arial">Arial</option>
              <option value="Calibri">Calibri</option>
              <option value="Georgia">Georgia</option>
              <option value="Helvetica">Helvetica</option>
            </select>
          </div>
          <div class="checkbox-wrapper">
            <input type="checkbox" id="includeToC" checked>
            <label for="includeToC">Include Table of Contents</label>
          </div>
        </div>

        <div id="sections-container"></div>

        <div class="button-container">
          <button class="secondary-button" onclick="google.script.host.close()">Cancel</button>
          <button onclick="submitSelections()">Create Structure</button>
        </div>

        <script>
          const presets = {
            standard: {
              name: "Standard Research Paper",
              info: "Traditional IMRAD format suitable for most research papers",
              sections: [
                {
                  title: "Introduction",
                  subsections: ["Background", "Problem Statement", "Research Objectives", "Research Questions", "Significance"]
                },
                {
                  title: "Methods",
                  subsections: ["Research Design", "Participants", "Data Collection", "Data Analysis", "Ethical Considerations"]
                },
                {
                  title: "Results",
                  subsections: ["Data Presentation", "Statistical Analysis", "Key Findings"]
                },
                {
                  title: "Discussion",
                  subsections: ["Interpretation", "Implications", "Limitations", "Future Research", "Conclusion"]
                }
              ]
            },
            systematic: {
              name: "Systematic Review",
              info: "Structured format for systematic literature reviews",
              sections: [
                {
                  title: "Introduction",
                  subsections: ["Background", "Review Objectives", "Research Questions"]
                },
                {
                  title: "Methods",
                  subsections: ["Search Strategy", "Inclusion Criteria", "Quality Assessment", "Data Extraction"]
                },
                {
                  title: "Results",
                  subsections: ["Study Selection", "Study Characteristics", "Quality Assessment", "Data Synthesis"]
                },
                {
                  title: "Discussion",
                  subsections: ["Summary of Evidence", "Limitations", "Conclusions", "Implications for Practice"]
                }
              ]
            },
            case: {
              name: "Case Study",
              info: "Detailed examination of a specific case or instance",
              sections: [
                {
                  title: "Introduction",
                  subsections: ["Background", "Case Overview", "Objectives"]
                },
                {
                  title: "Case Presentation",
                  subsections: ["Context", "Detailed Description", "Key Issues"]
                },
                {
                  title: "Methods",
                  subsections: ["Data Sources", "Analysis Approach"]
                },
                {
                  title: "Findings",
                  subsections: ["Key Observations", "Analysis Results"]
                },
                {
                  title: "Discussion",
                  subsections: ["Interpretation", "Lessons Learned", "Recommendations"]
                }
              ]
            },
            // Add more presets...
          };

          function loadPreset() {
            const selectedPreset = document.getElementById('presetType').value;
            const presetInfo = document.getElementById('presetInfo');
            
            if (selectedPreset === 'custom') {
              presetInfo.textContent = 'Create your own structure or load a saved template';
              return;
            }

            const preset = presets[selectedPreset];
            presetInfo.textContent = preset.info;
            renderSections(preset.sections);
          }

          function renderSections(sections) {
            const container = document.getElementById('sections-container');
            container.innerHTML = '';
            
            sections.forEach((section, index) => {
              const sectionDiv = document.createElement('div');
              sectionDiv.className = 'section';
              sectionDiv.innerHTML = \`
                <div class="section-title">
                  <input type="checkbox" checked>
                  <input type="text" value="\${section.title}" class="section-title-input">
                </div>
                <div class="subsections">
                  \${section.subsections.map(sub => \`
                    <div class="subsection">
                      <input type="checkbox" checked>
                      <input type="text" value="\${sub}">
                    </div>
                  \`).join('')}
                </div>
                <button onclick="addSubsection(this)">Add Subsection</button>
              \`;
              container.appendChild(sectionDiv);
            });
          }

          function showSaveDialog() {
            const templateName = prompt('Enter a name for this template:');
            if (templateName) {
              const template = {
                name: templateName,
                sections: getCurrentSections(),
                formatting: {
                  font: document.getElementById('fontFamily').value,
                  includeToC: document.getElementById('includeToC').checked
                }
              };
              google.script.run
                .withSuccessHandler(updateTemplatesList)
                .saveTemplate(template);
            }
          }

          function loadSavedTemplate() {
            const templateName = document.getElementById('savedTemplates').value;
            if (templateName) {
              google.script.run
                .withSuccessHandler(loadTemplate)
                .loadTemplate(templateName);
            }
          }

          function updateTemplatesList(templates) {
            const select = document.getElementById('savedTemplates');
            select.innerHTML = '<option value="">Select a saved template...</option>';
            templates.forEach(template => {
              const option = document.createElement('option');
              option.value = template.name;
              option.textContent = template.name;
              select.appendChild(option);
            });
          }

          // Load templates when page loads
          google.script.run
            .withSuccessHandler(updateTemplatesList)
            .getTemplates();

          // Initialize with standard preset
          loadPreset();
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(800)
    .setTitle('Configure IMRAD Structure');

  DocumentApp.getUi().showModalDialog(htmlOutput, 'Configure IMRAD Structure');
}

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

// Add these functions to handle template storage
function saveTemplate(template) {
  const userProperties = PropertiesService.getUserProperties();
  const templates = JSON.parse(userProperties.getProperty('templates') || '[]');
  templates.push(template);
  userProperties.setProperty('templates', JSON.stringify(templates));
  return templates;
}

function loadTemplate(templateName) {
  const userProperties = PropertiesService.getUserProperties();
  const templates = JSON.parse(userProperties.getProperty('templates') || '[]');
  return templates.find(t => t.name === templateName);
}

function getTemplates() {
  const userProperties = PropertiesService.getUserProperties();
  return JSON.parse(userProperties.getProperty('templates') || '[]');
}