# Resume/CV Markdown to ATS `.docx` Converter ⭐

A tool to convert your Markdown resume or cv into an ATS-friendly Word document that looks clean and professional while preserving your carefully crafted content.

## Overview 📚

This project allows you to maintain your resume in an easily editable Markdown format, then convert it to an ATS (Applicant Tracking System) optimized Word document with proper formatting for better parsing by job application systems.

🎨 **Your Markdown resume *must* use the same ("*Resume Markdown*") style as the [sample template](./sample/resume.sample.md)** 🎨

*(see the [Sample Template and Exmaple](#sample-template-and-example-) section for more details)*

## Key Features ⚡️

- Proper formatting of sections (contact, experience, education, etc.)
- Maintains hierarchy of job titles, companies, and dates
- Properly formats projects, skills, and responsibilities
- Creates an ATS-friendly document that parses well in applicant tracking systems

## Setup and Installation 📀

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

## Activation 🕹️

🌐 **Remember to *activate* the virtual environment *before* running any Python commands** 🌐

```bash
. .venv/bin/activate
```

> [!tip]
> Run `deactivate` to deactivate the *virtual environment*.

## Usage 👾

Convert your Markdown resume to a Word document.

> [!TIP]
> *The help screen can be accessed by running the following:*
> ```bash
> python resume_md_to_docx.py -h
> ```

### Basic usage 🐍

Run in interactive mode, prompting for inputs:

```bash
python resume_md_to_docx.py
```

Run in manual mode, specifying an input file:

> [!NOTE]
> This will create `My ATS Resume.docx` in the current directory.

```bash
python resume_md_to_docx.py -i resume.md
```

### Advanced usage 🐍

Specify an output filename:

```bash
python resume_md_to_docx.py -i resume.md -o your-filename.docx
```

Render heading levels as paragraphs instead of Word headings:

```bash
python resume_md_to_docx.py -i ~/Documents/resume.md -o ~/Desktop/custom-resume.docx -p h3 h4 h5 h6
```

> [!NOTE]
> The `-p` (or `--paragraph-headings`) option choices are: `h3`, `h4`, `h5`, `h6`

## Sample Template and Example 🤖

A [sample Markdown resume](./sample/resume.sample.md) (`sample/resume.sample.md`) is included in this project. You may copy or download it (removing `.sample` from the name if downloading) and use it as a template to create your own Markdown resume.

> [!IMPORTANT]
> For basic functionality, the **`h2`** level headings **should not** be changed; however if you feel so inclined, you can modify the `ResumeSection` *enum* according to your needs (see the [Resume Sections](#resume-sections-) section for more details).

You can [download the example `.docx` document](./sample/example.docx) (`sample/example.docx`) and open it in *Microsoft Word* or *Google Docs* (or another application capable of viewing `.docx` files) to see how the sample Markdown file is rendered.

## Resume Sections 🚀

The converter maps Markdown headings to ATS-friendly Word document headings using the `ResumeSection` enum. The **Markdown headings are *case-insensitive***. The default mappings are:

| Markdown Heading (h2) | Word Document Heading |
|----------------------|----------------------|
| About | PROFESSIONAL SUMMARY |
| Top Skills | CORE SKILLS |
| Experience | PROFESSIONAL EXPERIENCE |
| Education | EDUCATION |
| Linceces & Certifications | LICENSES & CERTIFICATIONS |
| Contact | CONTACT INFORMATION |

> [!Tip]
> If an `hr` (3 dashes, i.e. "---") is added immediately before a section (in your input `.md` file), that will put a page-break in the final document.

If you need to customize these mappings, you can modify the `ResumeSection` enum in [resume_md_to_docx.py](./resume_md_to_docx.py).

## Job Sub Sections 💼

Within job entries (particularly in the Experience section), various subsections can be used to structure your information. These are defined by the `JobSubsection` enum which maps markdown elements to properly formatted document sections. The **Markdown headings are *case-insensitive***. The default mappings are

| Markdown Element | Markdown Heading | Word Document Heading | Notes |
|------------------|--------------|---------------------|-------|
| h3 | highlights | Highlights | Used in the About section for key achievements |
| h5 | key skills | Technical Skills | Lists skills relevant to a specific role |
| h5 | summary | Summary | Brief overview of a position |
| h5 | internal | Internal | Internal project/responsibilities |
| h5 | project/client | Project/Client | Client project details |
| h6 | responsibilities overview | Responsibilities: | Project responsibilities |
| h6 | additional details | Additional Details: | Supplementary information |

These subsections help structure your job entries in a way that makes them more readable to both humans and ATS systems. For example, under each job, you might include a "Key Skills" subsection to highlight relevant technologies and abilities specific to that role.

## Basic Commands ⚙️

| Command | Description |
|---------|-------------|
| `make` | Alias for `make init` |
| `make help` | Show help information |
| `make list` | List all available commands |
| `make init` | Initialize the project |
| `make install` | Install dependencies |
| `make uninstall` | Uninstall dependencies |
| `make clean` | Clean up environment |

> [!NOTE]
> See the [Development](#development-) section (below) for advanced commands.

## Important Files 🗂️

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

> [!NOTE]
> The Makefile assumes a [POSIX compliant shell](https://wiki.archlinux.org/title/Command-line_shell) such as Bash, Zsh, or Dash.


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
