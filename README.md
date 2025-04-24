# Resume Markdown ➜ ATS-friendly Document ⭐

A tool to [convert your Markdown resume or cv into an ATS-friendly Word document](#usage-) that looks clean and professional while preserving your carefully crafted content. 🚀



## Overview 📚

This project allows you to maintain your resume in an easily editable Markdown format, then convert it to an ATS (Applicant Tracking System) optimized Word document with proper formatting for better parsing by job application systems.

🧬 **Your Markdown resume *must* use the same ("*Resume Markdown*") format as the [sample template](./sample/template/sample.md)** 🧬

*(see the [Sample Template](#sample-template-%EF%B8%8F) section for more details)*



## Key Features ⚡️

- Proper formatting of sections (contact, experience, education, etc.)
- Maintains hierarchy of job titles, companies, and dates
- Properly formats projects, skills, and responsibilities
- Creates an ATS-friendly document that parses well in applicant tracking systems



## Initial Setup 📀

> [!NOTE]
> *Your system needs to satisfy the [**system requirements**](#system-requirements-)*

The setup process involves running only 2 commands:

1. The `make` command creates any necessary pre-requisite files or directories, including creating a *virtual environment*, and ensures [`pip`](https://pip.pypa.io/en/stable/) is installed.
1. The `make install` command installs any required dependencies.

> [!WARNING]
> Although not strictly necessary, creating and [activating](#activation-️) a *virtual environment* is the **recommended** approach for most users. It causes the dependecies to be installed locally to this project, and not globally.

They can be run as separate commands, or as a single command, with the second dependent on the success of the first:

```bash
# Run as separate commands:
make
make install

# Run as a single command:
make && make install
```

> [!NOTE]
> See the [Basic Commands](#basic-commands-%EF%B8%8F) section for more commands.

> [!IMPORTANT]
> #### 🗒 Note about the `python` command
> All the commands in the [usage section](#usage-) assume [activation of a *virtual environment*](#activation-️), which, if created using the approach in this project (created with *python*) creates a `python` command alias. If you used an alternate setup approach and the `python` command isn't working, try `python3` instead. Or simply create an alias: **`alias python='python3'`**


### Activation 🕹️

🌐 **Remember to *activate* the virtual environment *before* running any Python commands** 🌐

```bash
. .venv/bin/activate
```

> [!TIP]
> Run `deactivate` to deactivate the *virtual environment*.



## Usage 👾

📘 **Convert your Markdown resume to a Word document (`.docx`)** 📘

*Please make sure the [initial setup](#initial-setup-) has been completed.*

> [!TIP]
> The help screen can be accessed by running the following:
> ```bash
> python src/resume_md_to_docx.py -h
> ```

> [!TIP]
> Spaces in file names can be escaped with a backslash (`\`), e.g. `path/to/my\ resume.md`

### Basic usage 🐍

By default, the name of the output file will match that of the input file, but with the appropriate extension. The **output files** will be in the project's [`output/`](./output/) directory unless other specified (with the `-o` or `--output` option).

#### ✨ Interactive mode 📱

By default, the command with no options or arguments, will cause the script to run in **interactive mode**, prompting the user (you) for inputs:

```bash
python src/resume_md_to_docx.py
```

#### Manual mode 🎛

Run in manual mode, specifying an input file:


```bash
# This will create a file called "resume.docx" in
# the "output/" directory, i.e. "output/resume.docx"
python src/resume_md_to_docx.py -i resume.md
```

Specify an output filename:

```bash
python src/resume_md_to_docx.py -i sample/example/example.md -o ~/Desktop/example\ ats\ resume.docx
```

> [!NOTE]
> If a `python: command not found` error occurs, see the [important note about the python command](#-note-about-the-python-command), in the usage section.


### ✨ Produce a PDF 📕

Adding `--pdf` to any of the above commands will also produce a `.pdf` file in the same directory as the `.docx` file (this will be the project's [`output/`](./output/) directory if the *output* option isn't set):

```bash
# This will create 2 files: "output/example.docx" and "output/example.pdf"
python src/resume_md_to_docx.py -i sample/example/example.md --pdf
```

> [!NOTE]
> The `--pdf` option isn't needed if running in *interactive mode*.


### All Options ⚙️

| Option | Long Form | Description | Default |
|--------|-----------|-------------|---------|
| `-c` | `--config` | Path to YAML configuration file | `resume_config.yaml` |
| `-h` | `--help` | Access the help screen | |
| `-i` | `--input` | Input markdown file | None (required in non-interactive mode) |
| `-o` | `--output` | Output Word document | `<input_file>.docx` in the output directory |
| `-p` | `--paragraph-headings` | Specify which heading levels to render as paragraphs instead of headings | None (all headings use Word styles) |
| `-I` | `--interactive` | Run in interactive mode, prompting for inputs | Auto-enabled when no other args provided |
| `-P` | `--pdf` | Also create a PDF version of the resume | Disabled |

> [!NOTE]
> The `-p` (or `--paragraph-headings`) option choices are: `h3`, `h4`, `h5`, `h6`. You can specify multiple heading levels by separating them with spaces (e.g. `<...command...> -p h5 h6`).

#### Examples 🤖

```bash
# Set input, output, and create a pdf
python src/resume_md_to_docx.py -i sample/example/example.md -o ~/Desktop/example\ ats\ resume.docx --pdf

# Set input, output, paragraph-headings, and create a pdf
python src/resume_md_to_docx.py -i sample/example/example.md -o ~/Desktop/example\ ats\ resume.docx -p h3 h4 h5 h6 --pdf

# Set input, output, paragraph-headings, create a pdf, and use a custom configuration file
python src/resume_md_to_docx.py -i sample/example/example.md -o ~/Desktop/example\ ats\ resume.docx --pdf -c custom_config.yaml
```



## Styling 🎨

A [configuration file](resume_config.yaml) (`resume_config.yaml`) is used to control certain stylings. It can be customized to modify how the `.docx` looks, to a limited degree.

> [!TIP]
> **One reason you might want to modify this file for your own purpose, is the font name** 🔠 (see below)

By default, `Helvetica Neue` is used as the base font. Your system should be able to figure out a compatible replacement automatically. However if you prefer to control the fonts, you can change the `font_name` property values:

```yaml
document_styles:
  Normal:
    font_name: "Arial"
    # ...

  Title:
    font_name: "Arial"
    # ...
```


## Sample Template 🖼️

A [sample Markdown resume](./sample/template/sample.md) (`sample/template/sample.md`) is included in this project. You may copy or download it and use it as a *template* to create your own Markdown resume.

> [!CAUTION]
> For basic functionality, the **`h2`** level headings **should not** be changed; however if you feel so inclined, you can modify the `ResumeSection` *enum* according to your needs (see the [Resume Sections](#resume-sections-) section for more details).

You can [download the sample `.docx` document](./sample/template/output/sample.docx) (`sample/template/output/sample.docx`) and open it in *Microsoft Word* or *Google Docs* (or another application capable of viewing `.docx` files) to see how the sample Markdown file is rendered.



## Example Resume ⚛️

֎ **An *"ai"* generated real-world [example](./sample/example/example.md) (`sample/example/example.md`) is also included in this project** ֎

⬇️ You can [download the example `.docx` document](./sample/example/output/example.docx) (`sample/example/output/example.docx`) and open it in a compatible application to see how the sample Markdown file is rendered.

👀 You can [view the example pdf](./sample/example/output/example.pdf) directly in your browser, if your browser supports it (most do).



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

### Modifications 🦾

If you need to customize these mappings, you can modify the `ResumeSection` enum in [src/resume_md_to_docx.py](./src/resume_md_to_docx.py).



## Job Sub Sections 💼

Within job entries (particularly in the Experience section), various subsections can be used to structure your information. These are defined by the `JobSubsection` enum which maps markdown elements to properly formatted document sections. The **Markdown headings are *case-insensitive***. The default mappings are:

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



## ✨ API Usage 🌀

The project includes a REST API that converts markdown resumes to ATS-friendly formats (DOCX and PDF). This allows you to integrate the conversion functionality into other applications or workflows.

### Starting the API Server 🚆

To start the API server:

```bash
# Using the make command
make api

# Or run directly
python src/api.py
```

By default, the server runs on `localhost:3000`. You can modify this in the api_config.yaml file.

### Configuration ⚙️

The API uses its own [configuration file](./api_config.yaml) (`api_config.yaml`) separate from the [resume styling configuration](#styling-).

### API Endpoints 🎸

1. #### Convert to DOCX 🦋

    Convert a markdown resume to DOCX format:

    ```
    POST /convert/docx
    POST /convert/docx/{filename}
    ```

1. #### Convert to PDF 🦋

    Convert a markdown resume to PDF format:

    ```
    POST /convert/pdf
    POST /convert/pdf/{filename}
    ```

### Request Parameters ⚙️

| Parameter | Description | Required |
|-----------|-------------|----------|
| `input_file` | Markdown resume file to convert | Yes |
| `paragraph_headings` | Heading levels to render as paragraphs (`h3`, `h4`, `h5`, `h6`) | No |
| `config_options` | JSON string with configuration overrides | No |

### Examples 🤖

#### Basic Conversion to DOCX 🦋

```bash
curl -X POST "http://localhost:3000/convert/docx" \
  -F "input_file=@resume.md" \
  -o resume_ats.docx
```

#### Convert to PDF with Paragraph Headings 🦋

```bash
curl -X POST "http://localhost:3000/convert/pdf" \
  -F "input_file=@resume.md" \
  -F "paragraph_headings=h3" \
  -F "paragraph_headings=h5" \
  -o resume_ats.pdf
```

#### Convert with Custom Configuration 🦋

```bash
curl -X POST "http://localhost:3000/convert/pdf" \
  -F "input_file=@resume.md" \
  -F "paragraph_headings=h5" \
  -F "config_options={\"style_constants\": {\"paragraph_lists\": true, \"column_layout\": true}}" \
  -o resume_ats.pdf
```

#### Using Filename in URL (better for curl -O) 🦋

```bash
curl -X POST "http://localhost:3000/convert/pdf/my_resume.pdf" \
  -F "input_file=@resume.md" \
  -O
```

This will download the file as `my_resume.pdf`.

### Swagger UI 🌊

The API includes Swagger documentation accessible at:

```
http://localhost:3000/swagger
```

This provides an interactive interface to:
- View all available endpoints
- Test API operations directly from the browser
- See detailed parameter and response documentation

### Error Responses 👻

Error responses are returned in JSON format:

```json
{
  "success": false,
  "message": "Error description"
}
```

> [!NOTE]
> HTTP status codes are appropriately used (400 for bad requests, 404 for not found, etc.).



## Important Files 🗂️

```
<project-root>/
├── output/                     # Default output directory
├── sample/
│   ├── example/
│   │   ├── example.md          # Real world example resume with mock data
│   │   └── output/
│   │       ├── example.docx    # Example docx ouput from example
│   │       └── example.pdf     # Example pdf ouput from example
│   └── template/
│       ├── sample.md           # Sample resume template
│       └── output/
│           ├── sample.docx     # Example docx ouput from sample
│           └── sample.pdf      # Example pdf ouput from sample
├── src/
│   ├── api.py                  # Main API script
│   └── resume_md_to_docx.py    # Main conversion script
├── Makefile                    # Contains helpful commands for managing the project
├── REAMDE.md                   # This README file
├── api_config.yaml             # The default api script configuration file
└── resume_config.yaml          # The default conversion script configuration file

```

> [!NOTE]
> *There are more files and directories in the project than what's shown above; the above just lists any files (and directories) that would be relevant to a typical user.*



## Basic Commands ⚙️

| Command | Description |
|---------|-------------|
| `make` | Alias for `make init` |
| `make api` | Run the flask app using the default configuration |
| `make help` | Show help information |
| `make list` | List all available commands |
| `make init` | Initialize the project |
| `make install` | Install dependencies |
| `make uninstall` | Uninstall dependencies |
| `make clean` | Clean up environment |

> [!NOTE]
> See the [Development](#development-) section (below) for advanced commands.



## System Requirements 🧰

- [Python 3.x](https://www.python.org/downloads/)
- [Make](https://www.gnu.org/software/make/)

> [!NOTE]
> The Makefile assumes a [POSIX compliant shell](https://wiki.archlinux.org/title/Command-line_shell) such as *Bash*, *Zsh*, or *Dash*.



## Development 🛠

For developers wishing to build this project:

| Command | Description |
|---------|-------------|
| `make install-dev` | Install development dependencies |
| `make uninstall-dev` | Uninstall development dependencies |
| `make build` | Rebuild `sample/template/output/sample.docx` from `sample/template/sample.md` |
| `make check` | Run linters without reformatting |
| `make lint` | Reformat code according to style guidelines |




# Contributing 💻

See [CONTRIBUTING.md](CONTRIBUTING.md) for information on contributing to this project.




# License 🪪

[![CC BY-NC-ND 4.0][cc-by-nc-nd-shield]][cc-by-nc-nd]

[This work](https://github.com/stairwaytowonderland/resume-md-to-ats-converter) © 2025 by [Andrew Haller](https://github.com/andrewhaller) is licensed under [Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International][cc-by-nc-nd].

[![CC BY-NC-ND 4.0][cc-by-nc-nd-image]][cc-by-nc-nd]

[cc-by-nc-nd]: http://creativecommons.org/licenses/by-nc-nd/4.0/
[cc-by-nc-nd-image]: https://licensebuttons.net/l/by-nc-nd/4.0/88x31.png
[cc-by-nc-nd-shield]: https://img.shields.io/badge/License-CC%20BY--NC--ND%204.0-lightgrey.svg
