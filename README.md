# PPTX to Markdown

This project provides a small Python CLI that converts a `.pptx` file into:

- a Markdown document with one section for each slide
- an assets folder containing exported slide images
- speaker notes extracted into dedicated Markdown sections

## Requirements

- Python 3.10+
- Dependencies from `requirements.txt`

## Installation

Open a terminal in this folder:

```powershell
Set-Location C:\REPOSITORY\PERSONALE\PPTXConvertToMd
python -m pip install -r requirements.txt
```

## Basic usage

Convert a PowerPoint file using the default output names:

```powershell
python .\pptx_to_md.py .\input.pptx
```

This generates:

- `input.md`
- `input_assets\`

## Custom output paths

You can choose both the Markdown file and the images folder:

```powershell
python .\pptx_to_md.py .\input.pptx -o .\output.md --images-dir .\output_assets
```

## Excluding speaker notes

Use `--no-notes` to omit the "Speaker notes" section from the output:

```powershell
python .\pptx_to_md.py .\input.pptx --no-notes
```

## Options

| Option | Short | Description |
| --- | --- | --- |
| `--output PATH` | `-o` | Path to the output `.md` file (default: same name as input) |
| `--images-dir PATH` | | Directory for extracted images (default: `<output_stem>_assets`) |
| `--no-notes` | | Exclude speaker notes from the output |

## What gets extracted

For each slide, the converter can include:

- text content
- tables converted to Markdown tables
- images exported as files and linked from Markdown
- speaker notes (can be excluded with `--no-notes`)

## Example

```powershell
python .\pptx_to_md.py C:\Slides\demo.pptx
```

Output:

- `C:\Slides\demo.md`
- `C:\Slides\demo_assets\`
