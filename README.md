
# Image to PPT Converter (img2ppt)

This skill allows Trae to convert architecture diagrams, flowcharts, and other structured images into fully editable PowerPoint slides.

## Installation

1.  Clone this repository or copy the `img2ppt` folder to your Trae skills directory:
    ```bash
    cp -r img2ppt ~/.trae/skills/
    ```

2.  Ensure you have the required Python packages installed in your environment:
    ```bash
    pip install python-pptx
    ```

## Usage

Simply drag an image into the Trae chat and ask:
> "Convert this image to an editable PPT."
> "把这张架构图转成 PPT。"

## How it Works

1.  **Analysis**: The AI vision model analyzes the image layout and elements.
2.  **Schema Generation**: It constructs a JSON representation of the diagram.
3.  **Rendering**: A Python script uses `python-pptx` to generate the native PowerPoint file.

## Contributing

Feel free to improve the prompt engineering in `SKILL.md` or suggest better Python rendering templates!
