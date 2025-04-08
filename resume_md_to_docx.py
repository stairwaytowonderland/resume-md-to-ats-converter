import argparse
import os
import re
import sys
from enum import Enum

import docx.oxml.shared
import markdown
from bs4 import BeautifulSoup
from bs4.element import PageElement as BS4_Element
from docx import Document
from docx.enum.text import WD_BREAK as DOCX_PAGE_BREAK
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT as DOCX_PARAGRAPH_ALIGN
from docx.opc.constants import RELATIONSHIP_TYPE as DOCX_REL
from docx.shared import Inches, Pt
from docx.text.paragraph import Paragraph as DOCX_Paragraph

##############################
# Define some defaults at module level for better performance
##############################
MD_LINK_PATTERN = re.compile(r"\[(.*?)\]\((.*?)\)")
URL_PATTERN = re.compile(r"https?://[^\s]+|www\.[^\s]+")
EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

DEFAULT_POINT_SIZE = 11
DEFAULT_INDENT_INCHES = 0.25
BULLET_INDENT_INCHES = 0.5


##############################
# Configs
##############################
class ResumeSection(Enum):
    """Maps markdown heading titles to their corresponding document headings

    Properties:
        markdown_heading (str): The text of the h2 heading in the markdown file
        docx_heading (str): The text to use as a heading in the Word document
        add_space_before_h3 (bool): Whether to add a blank line before each h3 heading
                                   (except the first after h2)
    """

    ABOUT = ("About", "PROFESSIONAL SUMMARY", False)
    SKILLS = ("Top Skills", "CORE SKILLS", False)
    EXPERIENCE = ("Experience", "PROFESSIONAL EXPERIENCE", True)
    EDUCATION = ("Education", "EDUCATION", True)
    CERTIFICATIONS = ("Licenses & certifications", "LICENSES & CERTIFICATIONS", True)
    CONTACT = ("Contact", "CONTACT INFORMATION", False)

    def __init__(
        self,
        markdown_heading: str,
        docx_heading: str,
        add_space_before_h3: bool = False,
    ):
        """Initialize ResumeSection enum

        Args:
            markdown_heading (str): The text of the h2 heading in the markdown file
            docx_heading (str): The text to use as a heading in the Word document
            add_space_before_h3 (bool): Whether to add a blank line before each h3 heading
                                       (except the first after h2). Defaults to False.
        """
        self.markdown_heading = markdown_heading
        self.docx_heading = docx_heading
        self.markdown_heading_lower = markdown_heading.lower()
        self.add_space_before_h3 = add_space_before_h3

    def matches(self, text):
        """Check if the given text matches this section's markdown_heading (case insensitive)

        Args:
            text (str): Text to compare against markdown_heading

        Returns:
            bool: True if text matches markdown_heading (case insensitive), False otherwise
        """
        return text.lower() == self.markdown_heading_lower


class JobSubsection(Enum):
    """Maps markdown subsection headings to their corresponding document headings

    Properties:
        markdown_heading_level (str): The HTML tag name (e.g., 'h5', 'h6') in markdown
        markdown_text_lower (str): The lowercase text content of the heading to match
        docx_heading (str): The text to use as a heading in the Word document
        separator (str): Optional separator to append after the heading (default: "")
        bold (bool): Whether the heading should be bold (default: True)
        italic (bool): Whether the heading should be italic (default: False)
    """

    KEY_SKILLS = ("h5", "key skills", "Technical Skills", "", True, False)
    SUMMARY = ("h5", "summary", "Summary", "", True, False)
    INTERNAL = ("h5", "internal", "Internal", "", True, False)
    PROJECT_CLIENT = ("h5", "project/client", "Project/Client", ": ", True, True)
    RESPONSIBILITIES = (
        "h6",
        "responsibilities overview",
        "Responsibilities",
        "",
        True,
        False,
    )
    ADDITIONAL_DETAILS = (
        "h6",
        "additional details",
        "Additional Details",
        "",
        True,
        False,
    )
    HIGHLIGHTS = ("h3", "highlights", "Highlights", "", True, False)

    def __init__(
        self,
        markdown_heading_level: str,
        markdown_text_lower: str,
        docx_heading: str,
        separator: str = "",
        bold: bool = True,
        italic: bool = False,
    ):
        """Initialize the JobSubsection with its properties"""
        self.markdown_heading_level = markdown_heading_level
        self.markdown_text_lower = markdown_text_lower
        self.docx_heading = docx_heading
        self.separator = separator
        self.bold = bold
        self.italic = italic

    @property
    def full_heading(self):
        """Return the full heading with separator

        Returns:
            str: The complete heading text with separator
        """
        return f"{self.docx_heading}{self.separator}"

    @classmethod
    def find_by_tag_and_text(cls, tag_name: str, text: str):
        """Find a JobSubsection by tag name and text content (case insensitive)

        Args:
            tag_name (str): HTML tag name to match (e.g., 'h5', 'h6')
            text (str): Text content to match (case insensitive)

        Returns:
            JobSubsection or None: The matching subsection or None if not found
        """
        text_lower = text.lower().strip()

        # First try exact match
        for subsection in cls:
            if (
                subsection.markdown_heading_level == tag_name
                and text_lower == subsection.markdown_text_lower
            ):
                return subsection

        # If no exact match, try partial match for elements
        # Create a dictionary of keywords to check against each subsection type
        if tag_name == "h6":  # Only do partial matching for h6 elements
            # Group subsections by heading level
            h6_subsections = [s for s in cls if s.markdown_heading_level == "h6"]

            # Try partial matching against each subsection's keywords
            for subsection in h6_subsections:
                # Extract the first word of the markdown_text_lower as the key keyword
                keyword = subsection.markdown_text_lower.split()[0]
                if keyword in text_lower:
                    return subsection

        return None


class MarkdownHeadingLevel(Enum):
    """Maps markdown heading levels to their corresponding document heading levels

    Properties:
        value (int): The Word document heading level (0-5) to use
        font_size (int): The font size in points for paragraph style formatting
    """

    H1 = (0, 16)  # Used for name at top
    H2 = (1, 14)  # Main section headings (About, Experience, etc.)
    H3 = (2, 14)  # Job titles
    H4 = (3, 12)  # Company names or role titles
    H5 = (4, 11)  # Subsections (Key Skills, Summary, etc.)
    H6 = (5, 10)  # Sub-subsections (Responsibilities, Additional Details)

    def __init__(self, value: str, font_size: str) -> None:
        """Initialize with heading level and font size

        Args:
            value (int): The Word document heading level (0-5) to use
            font_size (int): The font size in points for paragraph style formatting
        """
        self._value_ = value
        self.font_size = font_size

    @classmethod
    def get_level_for_tag(cls, tag_name: str) -> int | None:
        """Get the Word document heading level for a given tag name

        Args:
            tag_name (str): HTML tag name (e.g., 'h1', 'h2', etc.)

        Returns:
            int or None: The corresponding Word document heading level (0-5) or None if not found
        """
        if tag_name == "h1":
            return cls.H1.value
        elif tag_name == "h2":
            return cls.H2.value
        elif tag_name == "h3":
            return cls.H3.value
        elif tag_name == "h4":
            return cls.H4.value
        elif tag_name == "h5":
            return cls.H5.value
        elif tag_name == "h6":
            return cls.H6.value
        return None

    @classmethod
    def get_font_size_for_level(
        cls, heading_level: str, default_size: int = DEFAULT_POINT_SIZE
    ) -> int:
        """Get the font size for a given heading level

        Args:
            heading_level (int): The heading level (0-5)
            default_size (int): The default font size to return if not found (default: DEFAULT_POINT_SIZE (11))

        Returns:
            int: The font size in points
        """
        for level in cls:
            if level.value == heading_level:
                return level.font_size
        return default_size  # Default font size if not found


class PdfConverterPaths(Enum):
    """Enum to store paths to PDF converter executables on different platforms"""

    # Windows LibreOffice paths
    LIBREOFFICE_WINDOWS = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    # macOS LibreOffice path
    LIBREOFFICE_MACOS = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

    # Linux/Unix LibreOffice command name (assumed to be in PATH)
    LIBREOFFICE_LINUX = "libreoffice"


##############################
# Main Processors
##############################
def create_ats_resume(
    md_file: str, output_file: str, paragraph_style_headings: dict[str, bool] = None
) -> str:
    """Convert markdown resume to ATS-friendly Word document

    Args:
        md_file (str): Path to the markdown resume file
        output_file (str): Path where the output Word document will be saved
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags
                                                 ('h3', 'h4', etc.) to boolean values
                                                 indicating whether to use paragraph style
                                                 instead of heading style. Defaults to None.

    Returns:
        str: Path to the created document (DOCX or PDF if PDF creation succeeded)
    """
    # Default to using heading styles for all if not specified
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    # Read markdown file
    with open(md_file, "r") as file:
        md_content = file.read()

    # Convert markdown to HTML for easier parsing
    html = markdown.markdown(md_content)
    soup = BeautifulSoup(html, "html.parser")

    # Create document with standard margins
    document = Document()
    for section in document.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Define section processors with their specific processing functions
    section_processors = [
        (ResumeSection.ABOUT, process_header_section),
        (
            ResumeSection.ABOUT,
            lambda doc, soup, ps: process_about_section(doc, soup, ps),
        ),
        (
            ResumeSection.SKILLS,
            lambda doc, soup, ps: process_skills_section(doc, soup, ps),
        ),
        (
            ResumeSection.EXPERIENCE,
            lambda doc, soup, ps: process_experience_section(doc, soup, ps),
        ),
        (
            ResumeSection.EDUCATION,
            lambda doc, soup, ps: process_education_section(doc, soup, ps),
        ),
        (
            ResumeSection.CERTIFICATIONS,
            lambda doc, soup, ps: process_certifications_section(doc, soup, ps),
        ),
        (
            ResumeSection.CONTACT,
            lambda doc, soup, ps: process_contact_section(doc, soup, ps),
        ),
    ]

    # Process each section
    for section_type, processor in section_processors:
        processor(document, soup, paragraph_style_headings)

    # Save the document
    document.save(output_file)

    return output_file


def convert_to_pdf(docx_file: str, pdf_file: str = None):
    """Convert a DOCX file to PDF using available converters

    Args:
        docx_file (str): Path to the input DOCX file
        pdf_file (str, optional): Path for the output PDF file. If None,
                                 replaces the .docx extension with .pdf

    Returns:
        str: Path to the created PDF file, or None if conversion failed
    """
    if pdf_file is None:
        pdf_file = os.path.splitext(docx_file)[0] + ".pdf"

    # Try multiple conversion methods
    methods = [
        _convert_with_docx2pdf,
        _convert_with_libreoffice,
        _convert_with_win32com,
    ]

    for method in methods:
        try:
            if method(docx_file, pdf_file):
                return pdf_file
        except Exception as e:
            print(f"Could not convert using {method.__name__}: {str(e)}")

    print("❌ PDF conversion failed. Please install one of the following:")
    print("   - docx2pdf (pip install docx2pdf)")
    print("   - LibreOffice (https://www.libreoffice.org/)")
    print("   - Microsoft Word (Windows only)")
    return None


##############################
# Section Processors
##############################
def process_header_section(
    document: Document,
    soup: BeautifulSoup,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process the header (name and tagline) section

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    # Extract header (name)
    name = soup.find("h1").text

    # Add name as document title
    title = document.add_heading(name, MarkdownHeadingLevel.H1.value)
    title.alignment = DOCX_PARAGRAPH_ALIGN.CENTER

    # Add professional tagline if it exists - first paragraph after h1
    first_p = soup.find("h1").find_next_sibling()
    if first_p and first_p.name == "p":
        # Check if ANY paragraph headings are specified, which indicates preference for simpler styling
        use_paragraph_style = bool(
            paragraph_style_headings
        )  # True if any heading levels use paragraph style

        # Check if the paragraph contains emphasis (italics)
        em_tag = first_p.find("em")
        if em_tag:
            # Always use paragraph style with manual formatting when paragraph_style_headings is active
            if use_paragraph_style:
                tagline_para = document.add_paragraph()
                tagline_para.alignment = DOCX_PARAGRAPH_ALIGN.CENTER
                tagline_run = tagline_para.add_run(em_tag.text)
                tagline_run.italic = True

                # Set the font size to match heading style
                tagline_run.font.size = Pt(MarkdownHeadingLevel.H4.font_size)
            else:
                # Use Word's built-in Subtitle style
                tagline_para = document.add_paragraph(em_tag.text, style="Subtitle")
                tagline_para.alignment = DOCX_PARAGRAPH_ALIGN.CENTER
                # Keep it italic despite the style
                for run in tagline_para.runs:
                    run.italic = True

            # Add the rest of the first paragraph as a separate paragraph if it exists
            rest_of_p = first_p.text.replace(em_tag.text, "").strip()
            if rest_of_p:
                rest_para = document.add_paragraph()
                rest_para.alignment = DOCX_PARAGRAPH_ALIGN.CENTER
                rest_para.add_run(rest_of_p)
        else:
            # If no emphasis tag, just add the whole paragraph
            if use_paragraph_style:
                tagline_para = document.add_paragraph()
                tagline_para.alignment = DOCX_PARAGRAPH_ALIGN.CENTER
                tagline_para.add_run(first_p.text)
            else:
                # Use Word's built-in Subtitle style
                tagline_para = document.add_paragraph(first_p.text, style="Subtitle")
                tagline_para.alignment = DOCX_PARAGRAPH_ALIGN.CENTER

        # Check for additional paragraphs before the first h2 that might contain specialty areas
        current_p = first_p.find_next_sibling()
        # Simply process all paragraphs until we hit a non-paragraph element (like h2)
        while current_p and current_p.name == "p":
            specialty_para = document.add_paragraph()
            specialty_para.alignment = DOCX_PARAGRAPH_ALIGN.CENTER
            specialty_para.add_run(current_p.text)
            current_p = current_p.find_next_sibling()

    # Add horizontal line
    _add_horizontal_line_simple(document)


def process_about_section(
    document: Document,
    soup: BeautifulSoup,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process the About section

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    section_h2 = _prepare_section(
        document, soup, ResumeSection.ABOUT, paragraph_style_headings
    )

    if not section_h2:
        return

    current_element = section_h2.find_next_sibling()

    # Initialize processed_elements set if we're in the About section
    processed_elements = set()

    # Add all paragraphs until next h2
    while current_element and current_element.name != "h2":
        # Skip if already processed
        if current_element in processed_elements:
            current_element = current_element.find_next_sibling()
            continue

        # Check if this is a paragraph with strong element containing highlights
        highlights_subsection = None
        if (
            current_element.name == JobSubsection.HIGHLIGHTS.markdown_heading_level
            and current_element.text.strip().lower()
            == JobSubsection.HIGHLIGHTS.markdown_text_lower
        ):
            highlights_subsection = JobSubsection.HIGHLIGHTS

        # Handle regular paragraph
        if not highlights_subsection and current_element.name == "p":
            para = document.add_paragraph()

            # Process all elements of the paragraph to preserve formatting
            for child in current_element.children:
                # Check if this is a strong/bold element
                if getattr(child, "name", None) == "strong":
                    run = para.add_run(child.text)
                    run.bold = True
                # Check if this is an em/italic element
                elif getattr(child, "name", None) == "em":
                    run = para.add_run(child.text)
                    run.italic = True
                # Check if this is a link/anchor element
                elif getattr(child, "name", None) == "a" and child.get("href"):
                    _add_hyperlink(para, child.text, child.get("href"))
                # Otherwise, just add the text as-is
                else:
                    if child.string:
                        _process_text_for_hyperlinks(para, child.string)

            processed_elements.add(current_element)

        # Handle highlights subsection found in a paragraph or heading
        elif highlights_subsection:
            heading_level = (
                MarkdownHeadingLevel.get_level_for_tag(current_element.name)
                if current_element.name.startswith("h")
                else None
            )
            use_paragraph_style = paragraph_style_headings.get(
                current_element.name, False
            )
            _add_heading_or_paragraph(
                document,
                highlights_subsection.full_heading,
                heading_level,
                use_paragraph_style=use_paragraph_style,
                bold=highlights_subsection.bold,
                italic=highlights_subsection.italic,
            )
            processed_elements.add(current_element)

        # Handle bullet list
        elif current_element.name == "ul":
            _add_bullet_list(document, current_element)
            processed_elements.add(current_element)

        current_element = current_element.find_next_sibling()


def process_skills_section(
    document: Document,
    soup: BeautifulSoup,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process the Skills section

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    section_h2 = _prepare_section(
        document, soup, ResumeSection.SKILLS, paragraph_style_headings
    )

    if not section_h2:
        return

    current_element = section_h2.find_next_sibling()

    if current_element and current_element.name == "p":
        skills = [s.strip() for s in current_element.text.split("•")]
        skills = [s for s in skills if s]
        _add_formatted_paragraph(document, " | ".join(skills))


def process_experience_section(
    document: Document,
    soup: BeautifulSoup,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process the Experience section

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    section_h2 = _prepare_section(
        document, soup, ResumeSection.EXPERIENCE, paragraph_style_headings
    )

    if not section_h2:
        return

    # Find all job entries (h3 headings under Experience)
    current_element = section_h2.find_next_sibling()

    # Track for additional processing
    processed_elements = set()

    # Get the add_space_before_h3 setting from the section type
    add_space_before_h3 = ResumeSection.EXPERIENCE.add_space_before_h3

    while current_element and current_element.name != "h2":
        # Skip if already processed
        if current_element in processed_elements:
            current_element = current_element.find_next_sibling()
            continue

        # New job title (h3)
        if current_element.name == "h3":
            processed_elements = _process_job_entry(
                document,
                current_element,
                processed_elements,
                paragraph_style_headings,
                add_space_before_h3,
            )

        # Position titles under a company (h4)
        elif current_element.name == "h4" and current_element not in processed_elements:
            position_title = current_element.text.strip()

            # Check if this h4 is part of a company name after job title
            prev_h3 = current_element.find_previous_sibling("h3")
            next_after_h3 = prev_h3.find_next_sibling() if prev_h3 else None

            # Only process if this isn't a company name directly after job title
            if prev_h3 and next_after_h3 and next_after_h3 != current_element:
                # Use proper heading for position title
                heading_level = MarkdownHeadingLevel.get_level_for_tag(
                    current_element.name
                )
                use_paragraph_style = paragraph_style_headings.get(
                    current_element.name, False
                )
                _add_heading_or_paragraph(
                    document,
                    position_title,
                    heading_level,
                    use_paragraph_style=use_paragraph_style,
                )

                # Find date if available
                date_element = current_element.find_next_sibling()
                if (
                    date_element
                    and date_element.name == "p"
                    and date_element.find("em")
                ):
                    position_date = date_element.text.replace("*", "").strip()
                    date_para = document.add_paragraph()
                    date_run = date_para.add_run(position_date)
                    date_run.italic = True

                    # Mark as processed
                    processed_elements.add(date_element)

        elif current_element.name in ["h5", "h6"]:
            # Find matching subsection type
            subsection = JobSubsection.find_by_tag_and_text(
                current_element.name, current_element.text
            )

            if subsection:
                heading_level = MarkdownHeadingLevel.get_level_for_tag(
                    current_element.name
                )

                # PROJECT/CLIENT requires special handling with its own function
                if subsection == JobSubsection.PROJECT_CLIENT:
                    processed_elements = _process_project_section(
                        document,
                        current_element,
                        processed_elements,
                        paragraph_style_headings=paragraph_style_headings,
                    )

                # SUMMARY subsection
                elif subsection == JobSubsection.SUMMARY:
                    use_paragraph_style = paragraph_style_headings.get(
                        current_element.name, False
                    )
                    _add_heading_or_paragraph(
                        document,
                        subsection.full_heading,
                        heading_level,
                        use_paragraph_style=use_paragraph_style,
                        bold=subsection.bold,
                        italic=subsection.italic,
                    )

                # INTERNAL subsection
                elif subsection == JobSubsection.INTERNAL:
                    use_paragraph_style = paragraph_style_headings.get(
                        current_element.name, False
                    )
                    _add_heading_or_paragraph(
                        document,
                        subsection.full_heading,
                        heading_level,
                        use_paragraph_style=use_paragraph_style,
                        bold=subsection.bold,
                        italic=subsection.italic,
                    )

                    # Process elements under this internal section
                    next_element = current_element.find_next_sibling()
                    while next_element and next_element.name not in [
                        "h2",
                        "h3",
                        "h4",
                        "h5",
                    ]:
                        if next_element.name == "p":
                            # Process paragraph text
                            internal_para = document.add_paragraph()
                            _process_text_for_hyperlinks(
                                internal_para, next_element.text
                            )
                            processed_elements.add(next_element)
                        elif next_element.name == "ul":
                            for li in next_element.find_all("li"):
                                bullet_para = document.add_paragraph(
                                    style="List Bullet"
                                )
                                _left_indent_paragraph(
                                    bullet_para, BULLET_INDENT_INCHES
                                )  # Keep indentation for bullets
                                _process_text_for_hyperlinks(bullet_para, li.text)
                            processed_elements.add(next_element)

                        next_element = next_element.find_next_sibling()

                # KEY_SKILLS subsection
                elif subsection == JobSubsection.KEY_SKILLS:
                    use_paragraph_style = paragraph_style_headings.get(
                        current_element.name, False
                    )
                    skills_heading = _add_heading_or_paragraph(
                        document,
                        subsection.full_heading,
                        heading_level,
                        use_paragraph_style=use_paragraph_style,
                        bold=subsection.bold,
                        italic=subsection.italic,
                    )
                    skills_para = document.add_paragraph()

                    # Get skills from next element
                    next_element = current_element.find_next_sibling()
                    if next_element:
                        if next_element.name == "p":
                            skills_text = next_element.text.strip()
                            skills_para.add_run(skills_text)
                            processed_elements.add(next_element)
                        elif next_element.name == "ul":
                            skills_list = []
                            for li in next_element.find_all("li"):
                                skills_list.append(li.text.strip())
                            skills_para.add_run(" • ".join(skills_list))
                            processed_elements.add(next_element)

                    # Determine if we need to add a blank line after Key Skills
                    # We'll add a blank line if:
                    # 1. There are no more headings (end of section)
                    # 2. The next heading is an h3 (new job)
                    # 3. The next heading is an h4 (new role within same company)
                    # We won't add a blank line if there's an h5 or h6 after Key Skills
                    looking_ahead = current_element
                    next_heading = None

                    # Look for the next heading element
                    while looking_ahead and not next_heading:
                        looking_ahead = looking_ahead.find_next_sibling()
                        if looking_ahead and looking_ahead.name in [
                            "h3",
                            "h4",
                            "h5",
                            "h6",
                        ]:
                            next_heading = looking_ahead

                    # Add space if this is the last role or before a new role (most likely h4)
                    if not next_heading or next_heading.name in ["h4"]:
                        _add_space_paragraph(document, 8)

                # RESPONSIBILITIES subsection (standalone)
                elif (
                    subsection == JobSubsection.RESPONSIBILITIES
                    and current_element not in processed_elements
                ):
                    use_paragraph_style = paragraph_style_headings.get(
                        current_element.name, False
                    )
                    _add_heading_or_paragraph(
                        document,
                        subsection.full_heading,
                        heading_level,
                        use_paragraph_style=use_paragraph_style,
                        bold=subsection.bold,
                        italic=subsection.italic,
                    )

                    # Get content
                    next_element = current_element.find_next_sibling()
                    if next_element:
                        if next_element.name == "p":
                            resp_para = document.add_paragraph()
                            _process_text_for_hyperlinks(resp_para, next_element.text)
                            processed_elements.add(next_element)
                        elif next_element.name == "ul":
                            # Process bullet list
                            for li in next_element.find_all("li"):
                                bullet_para = document.add_paragraph(
                                    style="List Bullet"
                                )
                                _process_text_for_hyperlinks(bullet_para, li.text)
                            processed_elements.add(next_element)

                # ADDITIONAL_DETAILS subsection (standalone)
                elif (
                    subsection == JobSubsection.ADDITIONAL_DETAILS
                    and current_element not in processed_elements
                ):
                    use_paragraph_style = paragraph_style_headings.get(
                        current_element.name, False
                    )
                    _add_heading_or_paragraph(
                        document,
                        subsection.full_heading,
                        heading_level,
                        use_paragraph_style=use_paragraph_style,
                        bold=subsection.bold,
                        italic=subsection.italic,
                    )

                    # Get content (next element might be list items)
                    next_element = current_element.find_next_sibling()
                    if next_element and next_element.name == "ul":
                        for li in next_element.find_all("li"):
                            bullet_para = document.add_paragraph(style="List Bullet")

                            # Check if this bullet item contains a link
                            link = li.find("a")
                            if link and link.get("href"):
                                # Text before the link
                                prefix = ""
                                if link.previous_sibling:
                                    prefix = (
                                        link.previous_sibling.string
                                        if link.previous_sibling.string
                                        else ""
                                    )

                                # Text after the link
                                suffix = ""
                                if link.next_sibling:
                                    suffix = (
                                        link.next_sibling.string
                                        if link.next_sibling.string
                                        else ""
                                    )

                                # Add text before the link if any
                                if prefix.strip():
                                    bullet_para.add_run(prefix.strip())

                                # Add the hyperlink
                                _add_hyperlink(bullet_para, link.text, link.get("href"))

                                # Add text after the link if any
                                if suffix.strip():
                                    bullet_para.add_run(suffix.strip())
                            else:
                                # No HTML links, process for markdown links or plain text
                                _process_text_for_hyperlinks(bullet_para, li.text)

                        processed_elements.add(next_element)

        # Standalone bullet points
        elif current_element.name == "ul" and current_element not in processed_elements:
            for li in current_element.find_all("li"):
                bullet_para = document.add_paragraph(style="List Bullet")
                _process_text_for_hyperlinks(bullet_para, li.text)
            processed_elements.add(current_element)

        current_element = current_element.find_next_sibling()


def process_education_section(
    document: Document,
    soup: BeautifulSoup,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process the Education section

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    section_h2 = _prepare_section(
        document, soup, ResumeSection.EDUCATION, paragraph_style_headings
    )
    _process_simple_section(
        document,
        section_h2,
        add_space=ResumeSection.EDUCATION.add_space_before_h3,
        paragraph_style_headings=paragraph_style_headings,
    )


def process_certifications_section(
    document: Document,
    soup: BeautifulSoup,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process the Certifications section

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    section_h2 = _prepare_section(
        document, soup, ResumeSection.CERTIFICATIONS, paragraph_style_headings
    )
    _process_certifications(
        document,
        section_h2,
        paragraph_style_headings=paragraph_style_headings,
    )


def process_contact_section(
    document: Document,
    soup: BeautifulSoup,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process the Contact section

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    section_h2 = _prepare_section(
        document, soup, ResumeSection.CONTACT, paragraph_style_headings
    )
    _process_simple_section(
        document,
        section_h2,
        add_space=ResumeSection.CONTACT.add_space_before_h3,
        paragraph_style_headings=paragraph_style_headings,
    )


##############################
# Primary Helpers
##############################
def _prepare_section(
    document: Document,
    soup: BeautifulSoup,
    section_type: ResumeSection,
    paragraph_style_headings: dict[str, bool] = None,
) -> BS4_Element | None:
    """Universal preliminary section preparation

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        section_type: ResumeSection enum value
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        BeautifulSoup element or None: The section heading element if found, None otherwise
    """

    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    section_h2 = soup.find("h2", string=lambda text: section_type.matches(text))

    if not section_h2:
        return

    section_page_break = _has_hr_before_section(section_h2)

    # Add page break if requested
    if section_page_break:
        p = document.add_paragraph()
        run = p.add_run()
        run.add_break(DOCX_PAGE_BREAK.PAGE)

    # Add the section heading
    use_paragraph_style = paragraph_style_headings.get("h2", False)
    heading_level = MarkdownHeadingLevel.H2.value
    _add_heading_or_paragraph(
        document,
        section_type.docx_heading,
        heading_level,
        use_paragraph_style=use_paragraph_style,
    )

    return section_h2


def _process_simple_section(
    document: Document,
    section_h2: BS4_Element,
    add_space: bool = False,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process sections with simple paragraph-based content like Education and Contact.
    These sections typically have paragraphs with some bold (strong) elements.

    Args:
        document: The Word document object
        section_h2: The BeautifulSoup h2 element for the section
        add_space (bool, optional): Whether to add a space paragraph after the section. Defaults to False.
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    current_element = section_h2.find_next_sibling()

    while current_element and current_element.name != "h2":
        if current_element.name == "p":
            para = document.add_paragraph()
            for child in current_element.children:
                if getattr(child, "name", None) == "strong":
                    para.add_run(f"{child.text}: ").bold = True
                else:
                    if child.string and child.string.strip():
                        _process_text_for_hyperlinks(para, child.string.strip())
        elif current_element.name == "ul":
            # Handle bullet lists if they appear
            _add_bullet_list(document, current_element)

        current_element = current_element.find_next_sibling()

    # Add an extra space after the section if requested
    if add_space:
        _add_space_paragraph(document)


def _process_project_section(
    document: Document,
    project_element: BS4_Element,
    processed_elements: set[BS4_Element],
    paragraph_style_headings: dict[str, bool] = None,
) -> set[BS4_Element]:
    """Process a project/client section and its related elements

    Args:
        document: The Word document object
        project_element: BeautifulSoup element for the project heading
        processed_elements: Set of elements already processed
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        set: Updated set of processed elements
    """
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    # Get the next element to see if it contains the project details
    next_element = project_element.find_next_sibling()
    project_info = ""

    # If next element is a paragraph, it might contain the project name/duration
    if next_element and next_element.name == "p":
        project_info = next_element.text.strip()
        processed_elements.add(next_element)
        next_element = next_element.find_next_sibling()

    # Get the proper subsection
    subsection = JobSubsection.find_by_tag_and_text(
        project_element.name, project_element.text
    )
    if not subsection:
        subsection = JobSubsection.PROJECT_CLIENT  # Fallback

    heading_level = MarkdownHeadingLevel.get_level_for_tag(project_element.name)
    use_paragraph_style = paragraph_style_headings.get(project_element.name, False)

    # Prepare the project text
    project_text = subsection.full_heading
    if project_info:
        project_text += " " + project_info

    # Add the heading or formatted paragraph
    _add_heading_or_paragraph(
        document,
        project_text,
        heading_level,
        use_paragraph_style=use_paragraph_style,
        bold=subsection.bold,
        italic=subsection.italic,
    )

    # Process next elements under this project until another section
    while (
        next_element
        and next_element.name not in ["h2", "h3", "h4"]
        and not (next_element.name == "h5")
    ):
        # Find subsection for h6 elements
        h6_subsection = None
        if next_element.name == "h6":
            h6_subsection = JobSubsection.find_by_tag_and_text(
                next_element.name, next_element.text
            )

        # Responsibilities Overview
        if h6_subsection == JobSubsection.RESPONSIBILITIES:
            heading_level = MarkdownHeadingLevel.get_level_for_tag(next_element.name)
            use_paragraph_style = paragraph_style_headings.get(next_element.name, False)

            # Add the heading or paragraph
            if use_paragraph_style:
                resp_header = document.add_paragraph()
                _left_indent_paragraph(resp_header)  # Keep indentation
                resp_run = resp_header.add_run(h6_subsection.full_heading)
                resp_run.bold = h6_subsection.bold
                resp_run.italic = h6_subsection.italic
            else:
                resp_heading = document.add_heading(
                    h6_subsection.full_heading, level=heading_level
                )
                _left_indent_paragraph(resp_heading)  # Keep indentation

            # Get the paragraph with responsibilities
            resp_element = next_element.find_next_sibling()
            if resp_element and resp_element.name == "p":
                resp_para = document.add_paragraph()
                _process_text_for_hyperlinks(resp_para, resp_element.text)
                _left_indent_paragraph(resp_para)
                processed_elements.add(resp_element)

            processed_elements.add(next_element)

        # Additional Details
        elif h6_subsection == JobSubsection.ADDITIONAL_DETAILS:
            heading_level = MarkdownHeadingLevel.get_level_for_tag(next_element.name)
            use_paragraph_style = paragraph_style_headings.get(next_element.name, False)

            # Add the heading or paragraph
            if use_paragraph_style:
                details_header = document.add_paragraph()
                _left_indent_paragraph(details_header)  # Keep indentation
                details_run = details_header.add_run(h6_subsection.full_heading)
                details_run.bold = h6_subsection.bold
                details_run.italic = h6_subsection.italic
            else:
                details_heading = document.add_heading(
                    h6_subsection.full_heading, level=heading_level
                )
                _left_indent_paragraph(details_heading)  # Keep indentation

            processed_elements.add(next_element)

        # Bullet points
        elif next_element.name == "ul":
            for li in next_element.find_all("li"):
                bullet_para = document.add_paragraph(style="List Bullet")
                _left_indent_paragraph(
                    bullet_para, BULLET_INDENT_INCHES
                )  # Keep indentation for bullets
                _process_text_for_hyperlinks(bullet_para, li.text)
            processed_elements.add(next_element)

        next_element = next_element.find_next_sibling()

    return processed_elements


def _process_job_entry(
    document: Document,
    job_element: BS4_Element,
    processed_elements: set[BS4_Element],
    paragraph_style_headings: dict[str, bool] = None,
    add_space_before_h3: bool = False,
) -> set[BS4_Element]:
    """Process a job entry (h3) and its related elements

    Args:
        document: The Word document object
        job_element: BeautifulSoup element for the job heading
        processed_elements: Set of elements already processed
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values
        add_space_before_h3 (bool, optional): Whether to add a blank line before each h3 heading
                                           (except the first after h2). Defaults to False.

    Returns:
        set: Updated set of processed elements
    """
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    job_title = job_element.text.strip()

    # Check if this is NOT the first h3 after an h2
    # Only add space if the add_space_before_h3 parameter is True
    if add_space_before_h3:
        prev_heading = job_element.find_previous(["h2", "h3"])
        if prev_heading and prev_heading.name == "h3":
            # This is not the first h3 after h2, so add a blank line
            document.add_paragraph()

    # Check if h3 should use paragraph style
    use_paragraph_style = paragraph_style_headings.get("h3", False)
    heading_level = MarkdownHeadingLevel.get_level_for_tag(job_element.name)

    _add_heading_or_paragraph(
        document, job_title, heading_level, use_paragraph_style=use_paragraph_style
    )

    # Find company name (h4) if it exists
    next_element = job_element.find_next_sibling()

    # Add company if it exists
    if next_element and next_element.name == "h4":
        company_name = next_element.text.strip()

        # Check if h4 should use paragraph style
        use_paragraph_style = paragraph_style_headings.get("h4", False)
        company_heading_level = MarkdownHeadingLevel.get_level_for_tag(
            next_element.name
        )

        _add_heading_or_paragraph(
            document,
            company_name,
            company_heading_level,
            use_paragraph_style=use_paragraph_style,
        )

        # Mark as processed
        processed_elements.add(next_element)

        # Find date/location
        date_element = next_element.find_next_sibling()
        if date_element and date_element.name == "p" and date_element.find("em"):
            date_period = date_element.text.replace("*", "").strip()
            date_para = document.add_paragraph()
            date_run = date_para.add_run(date_period)
            date_run.italic = True

            # Mark as processed
            processed_elements.add(date_element)
    else:
        # Direct date under h3 (company header case)
        if next_element and next_element.name == "p" and next_element.find("em"):
            company_date = next_element.text.replace("*", "").strip()
            date_para = document.add_paragraph()
            date_run = date_para.add_run(company_date)
            date_run.italic = True

            # Mark as processed
            processed_elements.add(next_element)

    return processed_elements


def _add_heading_or_paragraph(
    document: Document,
    text: str,
    heading_level: int,
    use_paragraph_style: bool = False,
    bold: bool = True,
    italic: bool = False,
    font_size: int = None,
) -> DOCX_Paragraph:
    """Add either a heading or a formatted paragraph based on preference

    Args:
        document: The Word document object
        text (str): Text content for the heading/paragraph
        heading_level (int): The heading level (0-5) to use if not using paragraph style
        use_paragraph_style (bool): Whether to use paragraph style instead of heading style
        bold (bool): Whether to make the paragraph text bold (if using paragraph style)
        italic (bool): Whether to make the paragraph text italic (if using paragraph style)
        font_size (int, optional): Font size in points for paragraph style. Defaults to None.

    Returns:
        The created heading or paragraph object
    """
    if use_paragraph_style:
        para = document.add_paragraph()
        run = para.add_run(text)
        run.bold = bold
        run.italic = italic

        # Apply appropriate font size from the MarkdownHeadingLevel enum if not explicitly provided
        if font_size is None:
            size_pt = MarkdownHeadingLevel.get_font_size_for_level(heading_level)
            run.font.size = Pt(size_pt)
        else:
            run.font.size = Pt(font_size)

        return para
    else:
        return document.add_heading(text, level=heading_level)


def _process_certifications(
    document: Document,
    section_h2: BeautifulSoup,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process the certifications section with its specific structure

    Args:
        document: The Word document object
        section_h2: BeautifulSoup h2 element for the section
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    # Get the add_space_before_h3 setting from the section type
    add_space_before_h3 = ResumeSection.CERTIFICATIONS.add_space_before_h3

    current_element = section_h2.find_next_sibling()
    first_h3_after_h2 = True  # Track the first h3 after h2

    while current_element and current_element.name != "h2":
        # Process certification name (h3)
        if current_element.name == "h3":
            cert_name = current_element.text.strip()
            use_paragraph_style = paragraph_style_headings.get("h3", False)
            heading_level = MarkdownHeadingLevel.H3.value

            # Add blank line before h3 except for the first one
            # Only add space if add_space_before_h3 is True
            if add_space_before_h3 and not first_h3_after_h2:
                document.add_paragraph()
            else:
                first_h3_after_h2 = False  # After first h3 is processed

            _add_heading_or_paragraph(
                document,
                cert_name,
                heading_level,
                use_paragraph_style=use_paragraph_style,
            )

            # Look for next elements - either blockquote or organization info directly
            next_element = current_element.find_next_sibling()

            # Handle blockquote (optional)
            if next_element and next_element.name == "blockquote":
                # Process the blockquote contents
                _process_certification_blockquote(
                    document, next_element, paragraph_style_headings
                )

            # If no blockquote, look for organization info directly
            elif next_element:
                # Try to find organization info (could be bold text or heading)
                if next_element.name in ["h4", "h5", "h6", "p"] and next_element.find(
                    "strong"
                ):
                    # Extract organization text
                    org_text = next_element.find("strong").text.strip()
                    org_para = document.add_paragraph()
                    org_run = org_para.add_run(org_text)
                    org_run.bold = True

                    # Look for date information in the next element
                    date_element = next_element.find_next_sibling()
                    if date_element and date_element.find("em"):
                        em_tag = date_element.find("em")
                        date_text = em_tag.text.strip()
                        date_para = document.add_paragraph()

                        # Check if the date is hyperlinked
                        parent_a = em_tag.find_parent("a")
                        if parent_a and parent_a.get("href"):
                            _add_hyperlink(date_para, date_text, parent_a["href"])
                        else:
                            date_run = date_para.add_run(date_text)
                            date_run.italic = True

            # Add spacing after each certification
            # _add_space_paragraph(document)

        current_element = current_element.find_next_sibling()


def _process_certification_blockquote(
    document: Document,
    blockquote: BS4_Element,
    paragraph_style_headings: dict[str, bool] = None,
) -> None:
    """Process the contents of a certification blockquote

    Args:
        document: The Word document object
        blockquote: BeautifulSoup blockquote element containing certification details
        paragraph_style_headings (dict, optional): Dictionary mapping heading tags to boolean values

    Returns:
        None
    """
    if paragraph_style_headings is None:
        paragraph_style_headings = {}

    # Process the organization name (first element)
    org_element = blockquote.find(["h4", "h5", "h6", "strong"])
    if org_element:
        if org_element.name.startswith("h"):
            # It's a heading
            heading_level = MarkdownHeadingLevel.get_level_for_tag(org_element.name)
            use_paragraph_style = paragraph_style_headings.get(org_element.name, False)
            _add_heading_or_paragraph(
                document,
                org_element.text.strip(),
                heading_level,
                use_paragraph_style=use_paragraph_style,
            )
        elif org_element.name == "strong":
            # It's bold text
            org_para = document.add_paragraph()
            org_run = org_para.add_run(org_element.text.strip())
            org_run.bold = True

    # Process all paragraphs and text content
    # Find all p tags and loose text nodes directly under blockquote
    for item in blockquote.contents:
        # Skip elements we've already processed or that are empty
        if (isinstance(item, str) and not item.strip()) or item == org_element:
            continue

        # Check if it's a heading (h5, h6, etc.)
        if (
            hasattr(item, "name")
            and item.name
            and item.name.startswith("h")
            and item != org_element
        ):
            heading_level = MarkdownHeadingLevel.get_level_for_tag(item.name)
            use_paragraph_style = paragraph_style_headings.get(item.name, False)
            _add_heading_or_paragraph(
                document,
                item.text.strip(),
                heading_level,
                use_paragraph_style=use_paragraph_style,
            )
            continue

        # Check if it's the date with italics (should be last non-empty element)
        if hasattr(item, "find") and item.find("em"):
            em_tag = item.find("em")
            date_text = em_tag.text.strip()
            date_para = document.add_paragraph()

            # Check if inside a hyperlink
            parent_a = em_tag.find_parent("a")
            if parent_a and parent_a.get("href"):
                _add_hyperlink(date_para, date_text, parent_a["href"])
            else:
                date_run = date_para.add_run(date_text)
                date_run.italic = True
            continue

        # Handle normal text content (like "Some details")
        if isinstance(item, str) and item.strip():
            # It's a direct text node
            para = document.add_paragraph()
            _process_text_for_hyperlinks(para, item.strip())
        elif hasattr(item, "name"):
            if item.name == "p":
                # It's a paragraph
                para = document.add_paragraph()
                _process_text_for_hyperlinks(para, item.text.strip())
            elif item.name == "ul":
                # It's a bullet list
                _add_bullet_list(document, item)


##############################
# Inractive Mode Helper
##############################
def _run_interactive_mode() -> tuple[str, str, dict[str, bool], bool]:
    """Run in interactive mode, prompting the user for inputs

    Returns:
        tuple: (input_file, output_file, paragraph_style_headings)
    """
    print("\n🎯 Welcome to Resume Markdown to ATS Converter (Interactive Mode) 🎯\n")

    # Prompt for input file
    while True:
        input_file = input("📄 Enter the path to your Markdown resume file: ").strip()
        if not input_file:
            print("❌ Input file path cannot be empty. Please try again.")
            continue

        if not os.path.exists(input_file):
            print(f"❌ File '{input_file}' does not exist. Please enter a valid path.")
            continue

        break

    # Prompt for output file
    default_output = "My ATS Resume.docx"
    output_prompt = f"📝 Enter the output docx filename (default: '{default_output}'): "
    output_file = input(output_prompt).strip()
    if not output_file:
        output_file = default_output
        print(f"✅ Using default output: {output_file}")

    # Prompt for paragraph style headings
    print(
        "\n🔠 Choose heading levels to render as paragraphs instead of Word headings:"
    )
    print("   (Enter numbers separated by space, e.g., '3 4 5 6' for h3, h4, h5, h6)")
    print("   1. h3 - Job titles")
    print("   2. h4 - Company names")
    print("   3. h5 - Subsections (Key Skills, Summary, etc.)")
    print("   4. h6 - Sub-subsections (Responsibilities, Additional Details)")
    print("   0. None (use Word heading styles for all)")

    heading_choices = input(
        "👉 Your choices (e.g., '3 4 5 6' or '0' for none): "
    ).strip()

    paragraph_style_headings = {}
    if heading_choices != "0":
        chosen_numbers = [int(n) for n in heading_choices.split() if n.isdigit()]
        heading_map = {1: "h3", 2: "h4", 3: "h5", 4: "h6"}

        for num in chosen_numbers:
            if 1 <= num <= 4:
                heading_tag = heading_map[num]
                paragraph_style_headings[heading_tag] = True
                print(f"✅ Selected {heading_tag} for paragraph styling")
    else:
        print("✅ Using Word headings for all levels (no paragraph styling)")

    print("\n⚙️ Processing your resume...\n")

    create_pdf = (
        input("📄 Also create a PDF version? (y/n, default: n): ").strip().lower()
        == "y"
    )
    if create_pdf:
        print("✅ Will generate PDF output")

    return input_file, output_file, paragraph_style_headings, create_pdf


##############################
# PDF Helpers
##############################
def _convert_with_docx2pdf(docx_file: str, pdf_file: str) -> bool:
    """Convert using docx2pdf library"""
    try:
        from docx2pdf import convert

        convert(docx_file, pdf_file)
        return os.path.exists(pdf_file)
    except ImportError:
        return False


def _convert_with_libreoffice(docx_file: str, pdf_file: str) -> bool:
    """Convert using LibreOffice command line"""
    import platform
    import subprocess

    # Find the LibreOffice executable based on platform
    if platform.system() == "Windows":
        for path in PdfConverterPaths.LIBREOFFICE_WINDOWS.value:
            if os.path.exists(path):
                lo_exec = path
                break
        else:
            return False
    elif platform.system() == "Darwin":  # macOS
        lo_exec = PdfConverterPaths.LIBREOFFICE_MACOS.value
        if not os.path.exists(lo_exec):
            return False
    else:  # Linux/Unix
        lo_exec = PdfConverterPaths.LIBREOFFICE_LINUX.value  # Assume it's in PATH

    try:
        subprocess.run(
            [
                lo_exec,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                os.path.dirname(pdf_file) or ".",
                docx_file,
            ],
            check=True,
            stdout=subprocess.DEVNULL,
        )

        # LibreOffice saves to the original filename with .pdf extension
        temp_pdf = os.path.splitext(os.path.basename(docx_file))[0] + ".pdf"
        temp_pdf_path = os.path.join(os.path.dirname(pdf_file) or ".", temp_pdf)

        # If the output path is different from LibreOffice's default, move the file
        if os.path.abspath(temp_pdf_path) != os.path.abspath(pdf_file):
            os.rename(temp_pdf_path, pdf_file)

        return os.path.exists(pdf_file)
    except (subprocess.SubprocessError, OSError):
        return False


def _convert_with_win32com(docx_file: str, pdf_file: str) -> bool:
    """Convert using Microsoft Word COM automation (Windows only)"""
    import platform

    if platform.system() != "Windows":
        return False

    try:
        import win32com.client

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(os.path.abspath(docx_file))
        doc.SaveAs(os.path.abspath(pdf_file), FileFormat=17)  # 17 = PDF
        doc.Close()
        word.Quit()

        return os.path.exists(pdf_file)
    except ImportError:
        return False
    except Exception:
        # Clean up Word process if something went wrong
        try:
            doc.Close(False)
            word.Quit()
        except:
            pass
        return False


##############################
# Utilities
##############################
def _left_indent_paragraph(
    paragraph: DOCX_Paragraph, inches: float = DEFAULT_INDENT_INCHES
) -> DOCX_Paragraph:
    """Set the left indentation of a paragraph in inches

    Args:
        paragraph: The paragraph object to modify
        inches (float): Amount of indentation in inches

    Returns:
        paragraph: The modified paragraph object
    """
    # paragraph.paragraph_format._left_indent_paragraph = Inches(inches)
    paragraph.paragraph_format.left_indent = Inches(inches)
    return paragraph


def _add_bullet_list(
    document: Document, ul_element: BS4_Element, indentation: float = None
) -> DOCX_Paragraph:
    """Add bullet points from an unordered list element

    Args:
        document: The Word document object
        ul_element: BeautifulSoup element containing the unordered list
        indentation (float, optional): Left indentation in inches

    Returns:
        paragraph: The last bullet paragraph added
    """
    for li in ul_element.find_all("li"):
        bullet_para = document.add_paragraph(style="List Bullet")

        # Check if this bullet item contains a link
        link = li.find("a")
        if link and link.get("href"):
            # Text before the link
            prefix = ""
            if link.previous_sibling:
                prefix = (
                    link.previous_sibling.string if link.previous_sibling.string else ""
                )

            # Text after the link
            suffix = ""
            if link.next_sibling:
                suffix = link.next_sibling.string if link.next_sibling.string else ""

            # Add text before the link if any
            if prefix.strip():
                bullet_para.add_run(prefix.strip())

            # Add the hyperlink
            _add_hyperlink(bullet_para, link.text, link.get("href"))

            # Add text after the link if any
            if suffix.strip():
                bullet_para.add_run(suffix.strip())
        else:
            # No links, process as usual
            _process_text_for_hyperlinks(bullet_para, li.text)

        if indentation:
            _left_indent_paragraph(bullet_para, indentation)

    return bullet_para


def _add_formatted_paragraph(
    document: Document,
    text: str,
    bold: bool = False,
    italic: bool = False,
    alignment: DOCX_PARAGRAPH_ALIGN = None,
    indentation: float = None,
    font_size: int = None,
) -> DOCX_Paragraph:
    """Add a paragraph with consistent formatting

    Args:
        document: The Word document object
        text (str): Text content for the paragraph
        bold (bool, optional): Whether text should be bold. Defaults to False.
        italic (bool, optional): Whether text should be italic. Defaults to False.
        alignment (DOCX_PARAGRAPH_ALIGN, optional): Paragraph alignment. Defaults to None.
        indentation (float, optional): Left indentation in inches. Defaults to None.
        font_size (int, optional): Font size in points. Defaults to None.

    Returns:
        paragraph: The created paragraph object
    """
    para = document.add_paragraph()

    # Check if text contains URLs, emails, or markdown links
    is_link = _detect_link(text)[0]

    if bold or italic or not is_link:
        # If bold/italic formatting is needed or no links detected, use simple formatting
        run = para.add_run(text)
        run.bold = bold
        run.italic = italic
        if font_size:
            run.font.size = Pt(font_size)
    else:
        # Process for hyperlinks
        _process_text_for_hyperlinks(para, text)

    # Apply paragraph-level formatting
    if alignment:
        para.alignment = alignment
    if indentation:
        _left_indent_paragraph(para, indentation)

    return para


def _has_hr_before_section(section_h2: BS4_Element) -> bool:
    """Check if there's a horizontal rule (hr) element before a section heading

    Args:
        section_h2: BeautifulSoup element representing the section heading

    Returns:
        bool: True if there's an HR element immediately before this section, False otherwise
    """
    if not section_h2:
        return False

    prev_element = section_h2.previous_sibling
    # Skip whitespace text nodes
    while prev_element and isinstance(prev_element, str) and prev_element.strip() == "":
        prev_element = prev_element.previous_sibling

    # Check if the previous element is an HR
    return prev_element and prev_element.name == "hr"


def _detect_link(text: str) -> tuple[bool, str, str, str]:
    """Detect if text contains any kind of link (markdown link, URL, or email)

    Args:
        text (str): Text to check

    Returns:
        tuple: (is_link, display_text, url, matched_text)
            - is_link (bool): Whether any link was found
            - display_text (str): Text to display for the link
            - url (str): The URL for the hyperlink
            - matched_text (str): The full text that matched (for extraction)
    """
    link_types = [
        {
            "pattern": MD_LINK_PATTERN,
            "formatter": lambda m: (m.group(1), m.group(2), m.group(0)),
        },
        {
            "pattern": URL_PATTERN,
            "formatter": lambda m: (m.group(0), _format_url(m.group(0)), m.group(0)),
        },
        {
            "pattern": EMAIL_PATTERN,
            "formatter": lambda m: (m.group(0), f"mailto:{m.group(0)}", m.group(0)),
        },
    ]

    for link_type in link_types:
        match = link_type["pattern"].search(text)
        if match:
            display_text, url, matched_text = link_type["formatter"](match)
            return True, display_text, url, matched_text

    return False, text, "", ""


def _format_url(url: str) -> str:
    """Format URL to ensure it has proper scheme

    Args:
        url (str): URL to format

    Returns:
        str: Formatted URL with scheme
    """
    if url.startswith("www."):
        return "http://" + url
    return url


def _process_text_for_hyperlinks(paragraph: DOCX_Paragraph, text: str) -> None:
    """Process text to detect and add hyperlinks for Markdown links, URLs and email addresses

    Args:
        paragraph: The Word paragraph object to add content to
        text (str): Text to process for Markdown links, URLs and email addresses

    Returns:
        None: The paragraph is modified in place
    """
    # Check if text is None or empty
    if not text or not text.strip():
        return

    remaining_text = text

    # Keep finding links/URLs/emails until none remain
    while remaining_text:
        is_link, link_text, url, matched_text = _detect_link(remaining_text)

        if is_link:
            # Find the start position of the link
            link_position = remaining_text.find(matched_text)

            # Add text before the link (including any spaces)
            if link_position > 0:
                paragraph.add_run(remaining_text[:link_position])

            # Add the hyperlink with the appropriate text
            _add_hyperlink(paragraph, link_text, url)

            # Check if there's a space right after the link
            after_link_pos = link_position + len(matched_text)
            if (
                after_link_pos < len(remaining_text)
                and remaining_text[after_link_pos] == " "
            ):
                # Add the space separately to preserve it
                paragraph.add_run(" ")
                after_link_pos += 1

            # Continue with remaining text - after the full matched text
            remaining_text = remaining_text[after_link_pos:]
        else:
            # No more links, add the remaining text
            if remaining_text:
                paragraph.add_run(remaining_text)
            remaining_text = ""


def _add_hyperlink(
    paragraph: DOCX_Paragraph, text: str, url: str
) -> docx.oxml.shared.OxmlElement:
    """Add a hyperlink to a paragraph

    Args:
        paragraph: The paragraph to add the hyperlink to
        text (str): The text to display for the hyperlink
        url (str): The URL to link to

    Returns:
        Run: The created run object
    """
    # This gets access to the document
    part = paragraph.part
    # Create the relationship
    r_id = part.relate_to(url, DOCX_REL.HYPERLINK, is_external=True)

    # Create the hyperlink element
    hyperlink = docx.oxml.shared.OxmlElement("w:hyperlink")
    hyperlink.set(docx.oxml.shared.qn("r:id"), r_id)

    # Create a run inside the hyperlink
    new_run = docx.oxml.shared.OxmlElement("w:r")
    rPr = docx.oxml.shared.OxmlElement("w:rPr")

    # Add text style for hyperlinks (blue and underlined)
    color = docx.oxml.shared.OxmlElement("w:color")
    color.set(docx.oxml.shared.qn("w:val"), "0000FF")
    rPr.append(color)

    u = docx.oxml.shared.OxmlElement("w:u")
    u.set(docx.oxml.shared.qn("w:val"), "single")
    rPr.append(u)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    # Add the hyperlink to the paragraph
    paragraph._p.append(hyperlink)

    return hyperlink


def _add_horizontal_line_simple(document: Document) -> None:
    """Add a simple horizontal line to the document using underscores

    Args:
        document: The Word document object

    Returns:
        None
    """
    p = document.add_paragraph()
    p.alignment = DOCX_PARAGRAPH_ALIGN.CENTER
    p.add_run("_" * 50).bold = True

    # Add some space after the line
    # _add_space_paragraph(document)


def _add_space_paragraph(
    document: Document, font_size: int = DEFAULT_POINT_SIZE
) -> None:
    """Add a paragraph with extra space after it

    Args:
        document: The Word document object
        font_size (int): Controls both the spacing after paragraph and the font size of the run. Defaults to 12.

    Returns:
        None
    """
    p = document.add_paragraph()
    p.paragraph_format.space_after = Pt(font_size)


##############################
# Main Entry
##############################
if __name__ == "__main__":
    import os

    # Create detailed program description and examples
    program_description = """
    Convert a markdown resume to an ATS-friendly Word document.

    This script takes a markdown resume file and converts it to a properly formatted
    Word document that's optimized for Applicant Tracking Systems (ATS). It preserves
    the structure of your resume while ensuring proper formatting for employment history,
    skills, education, and other sections.
    """

    epilog_text = """
    Examples:
      python resume_md_to_docx.py
          - Runs in interactive mode (recommended for new users)

      python resume_md_to_docx.py -i resume.md
          - Converts resume.md to "My ATS Resume.docx"

      python resume_md_to_docx.py -i resume.md -o resume.docx
          - Converts resume.md to resume.docx

      python resume_md_to_docx.py -i resume.md --pdf
          - Converts resume.md to "My ATS Resume.docx" and "My ATS Resume.pdf"

      python resume_md_to_docx.py -i resume.md -o resume.docx --pdf
          - Converts resume.md to resume.docx and resume.pdf
    """

    # Parse command line arguments with enhanced help
    parser = argparse.ArgumentParser(
        description=program_description,
        epilog=epilog_text,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument("-i", "--input", dest="input_file", help="Input markdown file")
    parser.add_argument(
        "-o",
        "--output",
        dest="output_file",
        help='Output Word document (default: "My ATS Resume.docx")',
    )
    parser.add_argument(
        "-p",
        "--paragraph-headings",
        dest="paragraph_headings",
        nargs="+",
        choices=["h3", "h4", "h5", "h6"],
        help="Specify which heading levels to render as paragraphs instead of headings",
    )
    parser.add_argument(
        "-P",
        "--pdf",
        dest="create_pdf",
        action="store_true",
        help="Also create a PDF version of the resume",
    )
    parser.add_argument(
        "-I",
        "--interactive",
        dest="interactive",
        action="store_true",
        help="Run in interactive mode, prompting for inputs",
    )

    args = parser.parse_args()

    # Check if we should run in interactive mode
    # - No arguments were provided at all
    # - Only -I/--interactive flag was provided
    if (
        not (args.input_file or args.output_file or args.paragraph_headings)
        or args.interactive
    ):
        input_file, output_file, paragraph_style_headings, create_pdf = (
            _run_interactive_mode()
        )
    else:
        # Use command-line arguments
        input_file = args.input_file
        output_file = args.output_file or "My ATS Resume.docx"

        # Convert list to dictionary if provided
        paragraph_style_headings = {}
        if args.paragraph_headings:
            paragraph_style_headings = {h: True for h in args.paragraph_headings}

        # Show help if required arguments are missing
        if not input_file:
            parser.print_help()
            sys.exit(1)

        create_pdf = args.create_pdf

    # Use the arguments to create the resume
    result = create_ats_resume(
        input_file,
        output_file,
        paragraph_style_headings=paragraph_style_headings,
    )

    # Convert to PDF if requested
    if create_pdf:
        pdf_file = os.path.splitext(result)[0] + ".pdf"
        if convert_to_pdf(result, pdf_file):
            print(f"✅ Created PDF: {pdf_file}")

    print(f"🎉 ATS-friendly resume created: {result} 🎉")
