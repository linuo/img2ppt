
---
name: "img2ppt"
description: "Converts architecture diagrams or flowcharts (images) into editable PowerPoint slides. Invoke when the user provides an image and asks to convert it to PPT/slide."
---

# Image to PPT Converter

This skill converts images of diagrams (architecture, flowcharts, etc.) into editable PowerPoint (.pptx) slides.

## Capabilities

- Analyzes image content to identify shapes, text, colors, and layout structure.
- Generates a structured JSON representation of the diagram elements.
- Renders the JSON into a native .pptx file using Python code.
- Ensures all elements (shapes, text boxes, connectors) are editable in PowerPoint.

## Workflow

When the user provides an image and requests a PPT conversion:

1.  **Visual Analysis**:
    - Use your vision capabilities to analyze the image.
    - Identify all distinct visual elements: rectangles, rounded rectangles, cylinders (databases), arrows/connectors, and text labels.
    - Determine the approximate layout coordinates (relative to a 16:9 slide), sizes, and colors (RGB).
    - **CRITICAL**: Do NOT write Python code directly at this stage. Instead, construct a JSON object following the `DiagramSchema` below.

2.  **JSON Construction**:
    - Create a JSON object that strictly adheres to the schema.
    - Ensure all coordinates are normalized (e.g., assuming a 13.33 x 7.5 inch slide).
    - Map visual shapes to standard PPT shape types (e.g., `MSO_SHAPE.ROUNDED_RECTANGLE`).

3.  **Code Generation & Execution**:
    - Write a Python script that:
        - Imports `python-pptx`.
        - Defines the JSON data (paste the JSON you constructed).
        - Iterates through the JSON data to create slides and shapes.
        - Uses `python-pptx` to render each element with correct position, size, color, and text.
    - Execute the script to generate the `.pptx` file.

4.  **Delivery**:
    - Verify the file was created.
    - Inform the user that the file is ready for download/use.

## DiagramSchema (JSON)

Use this structure to represent the diagram:

```json
{
  "slide": {
    "width_inches": 13.33,
    "height_inches": 7.5,
    "background_color": "FFFFFF"
  },
  "elements": [
    {
      "type": "SHAPE | TEXT | CONNECTOR",
      "shape_type": "ROUNDED_RECTANGLE | RECTANGLE | OVAL | CYLINDER | ...", 
      "text": "Label Content",
      "position": {
        "x": 1.0, 
        "y": 2.0, 
        "w": 3.0, 
        "h": 1.0 
      },
      "style": {
        "fill_color": "FF0000", 
        "border_color": "000000",
        "border_width": 1.5,
        "is_dashed": false,
        "font_size": 12,
        "font_color": "000000",
        "alignment": "CENTER" 
      }
    },
    {
      "type": "CONNECTOR",
      "connector_type": "STRAIGHT | ELBOW | CURVED",
      "start": {"x": 1.0, "y": 2.5},
      "end": {"x": 4.0, "y": 2.5},
      "style": {
        "line_color": "000000",
        "line_width": 1.5,
        "start_arrow": false,
        "end_arrow": true
      }
    }
  ]
}
```

## Python Template (for Reference)

When generating the code, use this pattern to ensure reliability:

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ... (JSON Data Here) ...

def create_ppt(data):
    prs = Presentation()
    prs.slide_width = Inches(data['slide']['width_inches'])
    prs.slide_height = Inches(data['slide']['height_inches'])
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # ... (Rendering Logic) ...
    
    prs.save('output.pptx')
```
