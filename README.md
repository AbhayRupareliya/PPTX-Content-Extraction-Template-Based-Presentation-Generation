Over all system contain 6 Steps

1. Content Extraction
2. Template Analysis
3. Layout Detection
4. Build Embeddings / Vector Store
5. Template Selection (RAG + Scoring)
6. Final Presentation Generation

1. Content Extraction
   
The system first reads the input PowerPoint file and extracts all meaningful content and metadata.
Using python-pptx, the following information is extracted from each slide:
Slide title
Body text
Bullet points
Tables
Layout name
Shape position metadata

2. Template Analysis

All PPT templates stored in the templates folder are analyzed. For each template the system extracts:
Template name
Available slide layouts
Placeholder types
Placeholder count
Supported content structures

Step 3 — Layout Detection

The system analyzes the spatial arrangement of shapes on each slide.

Step 4 — Embedding Generation (Vector Store)

To improve template selection accuracy, the system converts both:
Slide structure
Template metadata into vector embeddings.

Step 5 — Template Selection

The system selects the most suitable template using a hybrid strategy.

Step 6 — Final Presentation Generation

After selecting the best template, the system rebuilds the presentation.




How the Scoring / Selection Works

The system selects the most suitable template using a hybrid selection strategy that combines:

1. Rule-Based Scoring
  Each template is evaluated against the extracted slide content using a scoring system.
  The score is calculated based on how well a template can support the slide structure.
  Factors Used for Scoring:
  1.1 Layout compatibility:	Whether the template layout matches the detected slide structure
  1.2 Placeholder capacity:	Whether the template has enough placeholders for the slide content
  1.3 Title support: Whether the template contains a title placeholder
  1.4 Table support:	Whether the template supports table content
3. Embedding similarity (RAG retrieval)
   To improve template selection accuracy, the system also uses embedding-based similarity.

This approach improves accuracy because it evaluates both structural compatibility and semantic similarity between slides and templates.




Why the chosen template was selected?

The template was selected because it best matched the structure and content of the input slides. The system first extracted slide titles, text elements, and layout information from the source presentation. It then analyzed all available templates to identify their layouts and placeholder capacities. Each template was scored based on how well it could support the extracted content, considering factors such as layout compatibility, number of placeholders, title support, and content structure. Additionally, embedding similarity (RAG-based retrieval) was used to compare slide structure with template metadata to find semantically similar layouts. The template with the highest overall compatibility score was chosen because it could most accurately represent the original slide content in the new presentation.



How to run my code?

1. Install required libraries
  Run the following command to install all dependencies:
  pip install python-pptx langchain langchain-openai faiss-cpu sentence-transformers
2. Prepare the project folders
  Place your files in the following structure:
      project

      input
         input.pptx
      
      templates
         template1.pptx
         template2.pptx
      
      output
      
      content_extraction.py
      template_analysis.py
      template_selection.py
      presentation_generation.py
      main.py

3. Add the input presentation

  Place the source PowerPoint file inside the input/ folder and name it: input/input.pptx

4. Add template presentations

  Place all candidate templates inside the templates/ folder.

5. Run the pipeline

  Execute the main script: python main.py

6. Check the generated output

    After execution, the system will generate a new presentation in: output/final_presentation.pptx


Library

Libraries Used:
The project uses python-pptx for reading and generating PowerPoint files, LangChain for building the AI-assisted pipeline, FAISS for vector similarity search in the RAG system, OpenAI embeddings for converting slide and template information into vectors, and sentence-transformers for generating embeddings used in template similarity comparison.
python-pptx – For reading, extracting content from, and generating PowerPoint (.pptx) files

LangChain – For building the AI-assisted workflow and integrating LLM-based reasoning
langchain-openai – For connecting LangChain with OpenAI models and embeddings
FAISS – For vector similarity search used in the RAG (Retrieval-Augmented Generation) pipeline
sentence-transformers – For generating embeddings to compare slide content with template structures
NumPy – For numerical operations and similarity calculations in embedding comparison
os (Python standard library) – For handling file and directory operations
