# Resume/CV Markdown to ATS `.docx` Converter ⭐

A tool to convert your Markdown resume or cv into an ATS-friendly Word document that looks clean and professional while preserving your carefully crafted content.

## Overview 📚

This project allows you to maintain your resume in an easily editable Markdown format, then convert it to an ATS (Applicant Tracking System) optimized Word document with proper formatting for better parsing by job application systems.

🎨 **Your Markdown resume *must* use the same ("*Resume Markdown*") style as our [sample template](./sample/resume.sample.md)** 🎨

*(see the [Sample Template and Exmaple](#sample-template-and-example-) section for more details)*

## Key Features ⚡️

- Proper formatting of sections (contact, experience, education, etc.)
- Maintains hierarchy of job titles, companies, and dates
- Properly formats projects, skills, and responsibilities
- Creates an ATS-friendly document that parses well in applicant tracking systems

## Installation 📀

Set up the project with:

```bash
make
```

Then install dependencies:

```bash
make install
```

> [!NOTE]
> See the [Makefile Commands (Basic)](#makefile-commands-basic-%EF%B8%8F) section (below) for more commands.

🌐 **Remember to *activate* the virtual environment *before* running any Python commands** 🌐

```bash
. .venv/bin/activate
```

> [!tip]
> Run `deactivate` to deactivate the *virtual environment*.

## Usage 👾

### Basic usage 🐍

Convert your Markdown resume to a Word document:

```bash
python resume_md_to_docx.py -i resume.md
```

This will create `My ATS Resume.docx` in the current directory.

### Advanced usage 🐍

Specify an output filename:

```bash
python resume_md_to_docx.py -i resume.md -o your-filename.docx
```

## Sample Template and Example 🤖

A [sample Markdown resume](./sample/resume.sample.md) (`sample/resume.sample.md`) is included in this project. You may copy or download it (removing `.sample` from the name if downloading) and use it as a template to create your own Markdown resume.

> [!IMPORTANT]
> For basic functionality, the **`h2`** level headings **should not** be changed; however if you feel so inclined, you can modify the `ResumeSection` *enum* according to your needs (see the [Resume Sections](#resume-sections-) section for more details).

You can [download the example `.docx` document](./sample/example.docx) (`sample/example.docx`) and open it in *Microsoft Word* or *Google Docs* (or another application capable of viewing `.docx` files) to see how the sample Markdown file is rendered.

## Resume Sections 🚀

The converter maps Markdown headings to ATS-friendly Word document headings using the `ResumeSection` enum. The default mappings are:

| Markdown Heading (h2) | Word Document Heading |
|----------------------|----------------------|
| About | PROFESSIONAL SUMMARY |
| Top Skills | CORE SKILLS |
| Experience | PROFESSIONAL EXPERIENCE |
| Education | EDUCATION |
| Contact | CONTACT INFORMATION |

If you need to customize these mappings, you can modify the `ResumeSection` enum in [resume_md_to_docx.py](./resume_md_to_docx.py).

## Makefile Commands (Basic) ⚙️

| Command | Description |
|---------|-------------|
| `make help` | Show help information |
| `make list` | List all available commands |
| `make init` | Initialize the project |
| `make install` | Install dependencies |
| `make uninstall` | Uninstall dependencies |
| `make clean` | Clean up environment |

> [!NOTE]
> See the [Development](#development-) section (below) for advanced commands.

## Project Structure 🗂️

```
<root>/
├── Makefile                    # Contains helpful commands for managing the project
├── resume_md_to_docx.py        # Main Python script for conversion
└── sample/
    ├── resume.sample.md        # Sample resume template
    └── example.docx            # Example docx ouput from sample with all stages
```

## Requirements ⚙️

- Python 3.x
- Make

## Development 🛠

For developers wishing to build this project:

| Command | Description |
|---------|-------------|
| `make install-dev` | Install development dependencies |
| `make uninstall-dev` | Uninstall development dependencies |
| `make build` | Rebuild `sample/example.docx` from `sample/resume.sample.md` |
| `make check` | Run linters without reformatting |
| `make lint` | Reformat code according to style guidelines |

# Contributing 💻

See [CONTRIBUTING.md](CONTRIBUTING.md) for information on contributing to this project.
