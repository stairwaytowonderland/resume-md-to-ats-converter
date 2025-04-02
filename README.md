# Resume/CV Markdown to ATS DOCX Converter

A tool to convert your Markdown resume or cv into an ATS-friendly Word document that looks clean and professional while preserving your carefully crafted content.

## Overview

This project allows you to maintain your resume in an easily editable Markdown format, then convert it to an ATS (Applicant Tracking System) optimized Word document with proper formatting for better parsing by job application systems.

## Installation

Set up the project with:

```bash
make
```

Then install dependencies:

```bash
make install
```

> [!NOTE]
> See the [Makefile Commands](#makefile-commands) section (below) for a full list of available commands.

Remember to activate the virtual environment before running any Python commands:

```bash
. .venv/bin/activate
```

## Usage

### Basic usage

Convert your Markdown resume to a Word document:

```bash
python resume_md_to_docx.py -i resume.md
```

This will create `My ATS Resume.docx` in the current directory.

### Advanced usage

Specify an output filename:

```bash
python resume_md_to_docx.py -i resume.md -o your-filename.docx
```

## Sample Template and Example

A sample Markdown resume -- `resume.sample.md` -- is included in this project. Please copy/paste it (removing `.sample` from the name) and use it as a guide to create your own resume.

> [!IMPORTANT]
> The **`h2`** level headings **must not** be changed.

You can download the sample `.docx` document -- `resume.sample.docx` -- and open it in Microsoft Word or Google Docs (or another application capable of viewing `.docx` files) to see how the sample Markdown file is rendered.

## Makefile Commands

| Command | Description |
|---------|-------------|
| `make help` | Show help information |
| `make list` | List all available commands |
| `make init` | Initialize the project |
| `make install` | Install dependencies |
| `make clean` | Clean up environment |
| `make test` | Run linters (without reformatting) |
| `make lint` | Run linters and reformat code |

## Project Structure

- `resume_md_to_docx.py` - Main Python script for conversion
- `resume.sample.md` - Sample resume template
- `Makefile` - Contains helpful commands for managing the project

## Features

- Proper formatting of sections (contact, experience, education, etc.)
- Maintains hierarchy of job titles, companies, and dates
- Properly formats projects, skills, and responsibilities
- Creates an ATS-friendly document that parses well in applicant tracking systems

## Requirements

- Python 3.x
- Make

## Development

For developers contributing to this project, use the linting tools:

```bash
make lint  # Reformat code according to style guidelines
make test  # Check code without reformatting
```
