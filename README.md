# PPTX-Content-Extraction-Template-Based-Presentation-Generation
Build a system (AI-assisted or Python-based) that can: 
1. Extract all meaningful content and metadata from a PPTX file.
2. Analyze multiple PPTX templates.
3. Select the most suitable template.
4. Recreate the presentation in the selected template.

Explanation

Step-1 Context Creation
  The system reads the input PPTX file and extracts meaningful content and metadata from each slide using python-pptx.
Extracted information includes:
Slide title
Body text
Bullet points
Table content (if present)
Slide layout name
Shape metadata (position, width, height)
  This information is stored in a structured JSON-like format, which becomes the intermediate representation of the presentation.

Step-2 Template Analysis

The system analyzes all templates available in the templates/ directory. For each template, the following information is extracted:

Template name
Available slide layouts
Placeholder types
Placeholder count
Layout structure

Step-3 Template Selection

The system selects the best template using a scoring-based selection algorithm.
Each template is evaluated against the extracted slide content.

Step-4 Final Presentation Generation

Once the best template is selected, the system reconstructs the presentation.
For each slide:
A new slide is created using the selected template.
The extracted title is inserted.
Body text and bullet points are inserted.
Tables are recreated if present.
