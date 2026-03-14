
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
    - **CRITICAL TEXT ANALYSIS**: Pay extreme attention to text styling. You MUST accurately estimate the relative font size for every piece of text (e.g., larger for titles, smaller for body text). You MUST determine if the text is bold.
    - **CANVAS BOUNDARY CHECK**: The slide has absolute dimensions (e.g., 13.33 x 7.5 inches). You MUST ensure that for EVERY element, `x + w` is strictly less than `slide.width_inches`, and `y + h` is strictly less than `slide.height_inches`. If your estimated coordinates exceed the boundaries, you must scale down the layout proportionately before generating the JSON. No elements should fall off the slide.
    - **CRITICAL**: Do NOT write Python code directly at this stage. Instead, construct a JSON object following the `DiagramSchema` below.

2.  **JSON Construction**:
    - Create a JSON object that strictly adheres to the schema.
    - Ensure all coordinates are normalized (e.g., assuming a 13.33 x 7.5 inch slide).
    - Map visual shapes to standard PPT shape types (e.g., `MSO_SHAPE.ROUNDED_RECTANGLE`).
    - **MANDATORY STYLE RULES**:
      - `font_color` MUST be explicitly defined for EVERY element containing text, regardless of the background color. Never assume a default color.
      - `font_size` MUST be explicitly defined for EVERY element containing text. Use your best judgment to reflect the visual hierarchy (e.g., titles might be 16-20, normal text 10-12).
      - `border_color` MUST be explicitly defined for ANY shape that visually has a border in the image (e.g., white boxes with gray outlines). Use `border_width` to control thickness.
      - `word_wrap`: Add `"word_wrap": false` in the `style` object for text that MUST remain on a single line (like short titles or labels). For paragraphs, use `"word_wrap": true`.
      - **CRITICAL WIDTH ESTIMATION**: For shapes containing text, ensure the width (`w`) is large enough to accommodate the text on a single line if `word_wrap` is false. Overestimating width slightly is better than unexpected line breaks.
      - **GLOBAL LAYOUT SCALING**: When analyzing a wide or complex diagram, calculate the total required width before assigning absolute X coordinates. If the diagram has many columns, allocate `w` values and `x` spacing such that the sum of the widest row fits within `slide.width_inches` (e.g., 13.33).

3.  **Code Generation & Execution**:
    - Do NOT write custom code. Use the existing skill workflow.
    - Save the JSON to a temporary file (e.g., `temp.json`).
    - Execute the provided renderer script: `python scripts/ppt_renderer.py temp.json output.pptx`.
    - Clean up the temporary JSON file after execution.

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
        "bold": false,
        "word_wrap": false,
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
