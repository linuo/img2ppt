
# Image to PPT Converter (img2ppt)

This skill enables AI agents to convert architecture diagrams, flowcharts, and other structured images into fully editable PowerPoint slides. It leverages the agent's multimodal capabilities (vision) to analyze the image and uses Python code to generate the native `.pptx` file.

## Capabilities

- **Visual Analysis**: Extracts layout, shapes, text, colors, and connectors from images.
- **Structured Output**: Maps visual elements to a standardized JSON schema.
- **Native Generation**: Renders editable PowerPoint slides using `python-pptx`.

## Installation

1.  Clone this repository or copy the `img2ppt` folder to your agent's skills directory (e.g., `.trae/skills/`, `.cursor/skills/`, or any custom agent workspace):
    ```bash
    git clone https://github.com/linuo/img2ppt.git
    ```

2.  Ensure you have the required Python packages installed in your environment:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

Simply drag an image into your AI agent's chat interface and ask:
> "Convert this image to an editable PPT."
> "把这张架构图转成可编辑的 PPT。"

## How it Works

1.  **Analysis (The Brain)**: The AI vision model analyzes the image layout and elements, constructing a JSON representation based on the `DiagramSchema` defined in `SKILL.md`.
2.  **Rendering (The Hand)**: The agent executes the Python script (`scripts/ppt_renderer.py`) which reads the JSON and uses `python-pptx` to generate the file.

## Requirements

- An AI agent with vision capabilities (e.g., Claude 3.5 Sonnet, GPT-4o).
- Python environment with `python-pptx` installed.

## Contributing

Feel free to improve the prompt engineering in `SKILL.md` or suggest better Python rendering templates!
