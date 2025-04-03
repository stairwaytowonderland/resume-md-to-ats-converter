# pip install python-docx markdown beautifulsoup4
# python resume_md_to_docx.py

import argparse
from enum import Enum

import docx.oxml.shared
from docx.opc.constants import RELATIONSHIP_TYPE

import markdown
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from docx.shared import Inches, Pt


class ResumeSection(Enum):
    """Maps markdown heading titles to their corresponding document headings

    Properties:
        markdown_heading (str): The text of the h2 heading in the markdown file
        docx_heading (str): The text to use as a heading in the Word document
    """
    ABOUT = ("About", "PROFESSIONAL SUMMARY")
    SKILLS = ("Top Skills", "CORE SKILLS")
    EXPERIENCE = ("Experience", "PROFESSIONAL EXPERIENCE")
    EDUCATION = ("Education", "EDUCATION")
    CERTIFICATIONS = ("Licenses & certifications", "LICENSES & CERTIFICATIONS")
    CONTACT = ("Contact", "CONTACT INFORMATION")

    def __init__(self, markdown_heading, docx_heading):
        """Initialize ResumeSection enum

        Args:
            markdown_heading (str): The text of the h2 heading in the markdown file
            docx_heading (str): The text to use as a heading in the Word document
        """
        self.markdown_heading = markdown_heading
        self.docx_heading = docx_heading
        self.markdown_heading_lower = markdown_heading.lower()

    def matches(self, text):
        """Check if the given text matches this section's markdown_heading (case insensitive)

        Args:
            text (str): Text to compare against markdown_heading

        Returns:
            bool: True if text matches markdown_heading (case insensitive), False otherwise
        """
        return text.lower() == self.markdown_heading_lower

    @classmethod
    def find_by_text(cls, text):
        """Find ResumeSection by heading text (case insensitive)

        Args:
            text (str): Text to search for

        Returns:
            ResumeSection or None: Matching section or None if not found
        """
        if not text:
            return None
        text_lower = text.lower()
        for section in cls:
            if section.markdown_heading_lower == text_lower:
                return section
        return None


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
    PROJECT_CLIENT = ("h5", "project/client", "Project/Client", "", True, True)
    RESPONSIBILITIES = (
        "h6",
        "responsibilities overview",
        "Responsibilities",
        ":",
        True,
        False,
    )
    ADDITIONAL_DETAILS = (
        "h6",
        "additional details",
        "Additional Details",
        ":",
        True,
        False,
    )
    HIGHLIGHTS = ("h3", "highlights", "Highlights", "", True, False)

    def __init__(
        self,
        markdown_heading_level,
        markdown_text_lower,
        docx_heading,
        separator="",
        bold=True,
        italic=False,
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
    def find_by_tag_and_text(cls, tag_name, text):
        """Find a JobSubsection by tag name and text content (case insensitive)

        Args:
            tag_name (str): HTML tag name to match (e.g., 'h5', 'h6')
            text (str): Text content to match (case insensitive)

        Returns:
            JobSubsection or None: The matching subsection or None if not found
        """
        text_lower = text.lower().strip()
        for subsection in cls:
            if (
                subsection.markdown_heading_level == tag_name
                and text_lower == subsection.markdown_text_lower
            ):
                return subsection
        return None


class MarkdownHeadingLevel(Enum):
    """Maps markdown heading levels to their corresponding document heading levels

    Properties:
        value (int): The Word document heading level (0-5) to use
    """

    H1 = 0  # Used for name at top
    H2 = 1  # Main section headings (About, Experience, etc.)
    H3 = 2  # Job titles
    H4 = 3  # Company names or role titles
    H5 = 4  # Subsections (Key Skills, Summary, etc.)
    H6 = 5  # Sub-subsections (Responsibilities, Additional Details)

    @classmethod
    def get_level_for_tag(cls, tag_name):
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


def create_ats_resume(md_file, output_file):
    """Convert markdown resume to ATS-friendly Word document

    Args:
        md_file (str): Path to the markdown resume file
        output_file (str): Path where the output Word document will be saved

    Returns:
        str: Path to the created document
    """
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

    # Extract header (name)
    name = soup.find("h1").text

    # Add name as document title
    title = document.add_heading(name, MarkdownHeadingLevel.H1.value)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Add professional tagline if it exists - first paragraph after h1
    first_p = soup.find("h1").find_next_sibling()
    if first_p and first_p.name == "p":
        # Check if the paragraph contains emphasis (italics)
        em_tag = first_p.find("em")
        if em_tag:
            tagline_para = document.add_paragraph()
            tagline_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            tagline_run = tagline_para.add_run(em_tag.text)
            tagline_run.italic = True

            # Add the rest of the first paragraph as a separate paragraph if it exists
            rest_of_p = first_p.text.replace(em_tag.text, "").strip()
            if rest_of_p:
                rest_para = document.add_paragraph()
                rest_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                rest_para.add_run(rest_of_p)
        else:
            # If no emphasis tag, just add the whole paragraph
            tagline_para = document.add_paragraph()
            tagline_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            tagline_para.add_run(first_p.text)

        # Check for additional paragraphs before the first h2 that might contain specialty areas
        current_p = first_p.find_next_sibling()
        # Simply process all paragraphs until we hit a non-paragraph element (like h2)
        while current_p and current_p.name == "p":
            specialty_para = document.add_paragraph()
            specialty_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            specialty_para.add_run(current_p.text)
            current_p = current_p.find_next_sibling()

    # Add horizontal line
    add_horizontal_line_simple(document)

    # Process About section
    about_h2 = soup.find("h2", string=lambda text: ResumeSection.ABOUT.matches(text))
    about_page_break = has_hr_before_section(about_h2)
    about_h2 = process_section(document, soup, ResumeSection.ABOUT, page_break=about_page_break)
    if about_h2:
        current_p = about_h2.find_next_sibling()

        # Initialize processed_elements set if we're in the About section
        processed_elements = set()

        # Add all paragraphs until next h2
        while current_p and current_p.name != "h2":
            # Skip if already processed
            if current_p in processed_elements:
                current_p = current_p.find_next_sibling()
                continue

            # Check if this is a paragraph with strong element containing highlights
            highlights_subsection = None
            if (
                current_p.name == JobSubsection.HIGHLIGHTS.markdown_heading_level
                and current_p.text.strip().lower()
                == JobSubsection.HIGHLIGHTS.markdown_text_lower
            ):
                highlights_subsection = JobSubsection.HIGHLIGHTS

            # Handle regular paragraph
            if not highlights_subsection and current_p.name == "p":
                para = document.add_paragraph()

                # Process all elements of the paragraph to preserve formatting
                for child in current_p.children:
                    # Check if this is a strong/bold element
                    if getattr(child, "name", None) == "strong":
                        run = para.add_run(child.text)
                        run.bold = True
                    # Check if this is an em/italic element
                    elif getattr(child, "name", None) == "em":
                        run = para.add_run(child.text)
                        run.italic = True
                    # Otherwise, just add the text as-is
                    else:
                        if child.string and child.string.strip():
                            para.add_run(child.string.strip())

                processed_elements.add(current_p)

            # Handle highlights subsection found in a paragraph or heading
            elif highlights_subsection:
                heading_level = (
                    MarkdownHeadingLevel.get_level_for_tag(current_p.name)
                    if current_p.name.startswith("h")
                    else None
                )
                if heading_level is not None:
                    # Use heading style
                    highlights_heading = document.add_heading(
                        highlights_subsection.full_heading, level=heading_level
                    )
                else:
                    # Otherwise use paragraph with manual formatting
                    highlights_para = add_formatted_paragraph(
                        document,
                        highlights_subsection.full_heading,
                        bold=highlights_subsection.bold,
                        italic=highlights_subsection.italic,
                    )
                processed_elements.add(current_p)

            # Handle bullet list
            elif current_p.name == "ul":
                add_bullet_list(document, current_p)
                processed_elements.add(current_p)

            current_p = current_p.find_next_sibling()

    # Process Top Skills section
    skills_h2 = soup.find("h2", string=lambda text: ResumeSection.SKILLS.matches(text))
    skills_page_break = has_hr_before_section(skills_h2)
    skills_h2 = process_section(document, soup, ResumeSection.SKILLS, page_break=skills_page_break)
    if skills_h2:
        current_p = skills_h2.find_next_sibling()
        if current_p and current_p.name == "p":
            skills = [s.strip() for s in current_p.text.split("•")]
            skills = [s for s in skills if s]
            add_formatted_paragraph(document, " | ".join(skills))

    # Process Experience section
    experience_h2 = soup.find("h2", string=lambda text: ResumeSection.EXPERIENCE.matches(text))
    experience_page_break = has_hr_before_section(experience_h2)
    if experience_h2:
        # Use page break if there's an HR before this section
        if experience_page_break:
            p = document.add_paragraph()
            run = p.add_run()
            run.add_break(docx.enum.text.WD_BREAK.PAGE)

        document.add_heading(
            ResumeSection.EXPERIENCE.docx_heading, level=MarkdownHeadingLevel.H2.value
        )

        # Find all job entries (h3 headings under Experience)
        current_element = experience_h2.find_next_sibling()

        # Track for additional processing
        processed_elements = set()

        while current_element and current_element.name != "h2":
            # Skip if already processed
            if current_element in processed_elements:
                current_element = current_element.find_next_sibling()
                continue

            # New job title (h3)
            if current_element.name == "h3":
                processed_elements = process_job_entry(
                    document, current_element, processed_elements
                )

            # Position titles under a company (h4)
            elif (
                current_element.name == "h4"
                and current_element not in processed_elements
            ):
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
                    document.add_heading(position_title, level=heading_level)

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
                        processed_elements = process_project_section(
                            document, current_element, processed_elements
                        )

                    # SUMMARY subsection
                    elif subsection == JobSubsection.SUMMARY:
                        if heading_level is not None:
                            # Use heading style
                            summary_heading = document.add_heading(
                                subsection.full_heading, level=heading_level
                            )
                        else:
                            # Otherwise use paragraph with manual formatting
                            summary_para = document.add_paragraph()
                            summary_run = summary_para.add_run(subsection.full_heading)
                            summary_run.bold = subsection.bold
                            summary_run.italic = subsection.italic

                    # INTERNAL subsection
                    elif subsection == JobSubsection.INTERNAL:
                        if heading_level is not None:
                            # Use heading style
                            internal_heading = document.add_heading(
                                subsection.full_heading, level=heading_level
                            )
                        else:
                            # Otherwise use paragraph with manual formatting
                            internal_para = document.add_paragraph()
                            internal_run = internal_para.add_run(
                                subsection.full_heading
                            )
                            internal_run.bold = subsection.bold
                            internal_run.italic = subsection.italic

                        # Process elements under this internal section
                        next_element = current_element.find_next_sibling()
                        while next_element and next_element.name not in [
                            "h2",
                            "h3",
                            "h4",
                            "h5",
                        ]:
                            if next_element.name == "ul":
                                for li in next_element.find_all("li"):
                                    bullet_para = document.add_paragraph(
                                        style="List Bullet"
                                    )
                                    left_indent_paragraph(
                                        bullet_para, 0.5
                                    )  # Keep indentation for bullets
                                    bullet_para.add_run(li.text)
                                processed_elements.add(next_element)

                            next_element = next_element.find_next_sibling()

                    # KEY_SKILLS subsection
                    elif subsection == JobSubsection.KEY_SKILLS:
                        if heading_level is not None:
                            # Use heading style
                            skills_heading = document.add_heading(
                                subsection.full_heading, level=heading_level
                            )
                            skills_para = document.add_paragraph()
                        else:
                            # Otherwise use paragraph with manual formatting
                            skills_para = document.add_paragraph()
                            skills_run = skills_para.add_run(subsection.full_heading)
                            skills_run.bold = subsection.bold
                            skills_run.italic = subsection.italic

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

                        document.add_paragraph()  # Add spacing

                    # RESPONSIBILITIES subsection (standalone)
                    elif (
                        subsection == JobSubsection.RESPONSIBILITIES
                        and current_element not in processed_elements
                    ):
                        if heading_level is not None:
                            # Use heading style
                            resp_heading = document.add_heading(
                                subsection.full_heading, level=heading_level
                            )
                        else:
                            # Otherwise use paragraph with manual formatting
                            resp_header = document.add_paragraph()
                            resp_run = resp_header.add_run(subsection.full_heading)
                            resp_run.bold = subsection.bold
                            resp_run.italic = subsection.italic

                        # Get content
                        next_element = current_element.find_next_sibling()
                        if next_element and next_element.name == "p":
                            document.add_paragraph(next_element.text)
                            processed_elements.add(next_element)

                    # ADDITIONAL_DETAILS subsection (standalone)
                    elif (
                        subsection == JobSubsection.ADDITIONAL_DETAILS
                        and current_element not in processed_elements
                    ):
                        if heading_level is not None:
                            # Use heading style
                            details_heading = document.add_heading(
                                subsection.full_heading, level=heading_level
                            )
                        else:
                            # Otherwise use paragraph with manual formatting
                            details_header = document.add_paragraph()
                            details_run = details_header.add_run(
                                subsection.full_heading
                            )
                            details_run.bold = subsection.bold
                            details_run.italic = subsection.italic

                        # Get content (next element might be list items)
                        next_element = current_element.find_next_sibling()
                        if next_element and next_element.name == "ul":
                            add_bullet_list(document, next_element)
                            processed_elements.add(next_element)

            # Standalone bullet points
            elif (
                current_element.name == "ul"
                and current_element not in processed_elements
            ):
                for li in current_element.find_all("li"):
                    bullet_para = document.add_paragraph(style="List Bullet")
                    bullet_para.add_run(li.text)

            current_element = current_element.find_next_sibling()

    # Process Education section
    education_h2 = soup.find("h2", string=lambda text: ResumeSection.EDUCATION.matches(text))
    education_page_break = has_hr_before_section(education_h2)
    process_simple_section(document, soup, ResumeSection.EDUCATION, add_space=True,
                        page_break=education_page_break)

    # Process Licenses & certifications section
    certifications_h2 = soup.find("h2", string=lambda text: ResumeSection.CERTIFICATIONS.matches(text))
    certifications_page_break = has_hr_before_section(certifications_h2)
    process_certifications(document, soup, ResumeSection.CERTIFICATIONS,
                        page_break=certifications_page_break)

    # Add contact information
    contact_h2 = soup.find("h2", string=lambda text: ResumeSection.CONTACT.matches(text))
    contact_page_break = has_hr_before_section(contact_h2)
    process_simple_section(document, soup, ResumeSection.CONTACT,
                        page_break=contact_page_break)

    # Save the document
    document.save(output_file)
    return output_file


def left_indent_paragraph(paragraph, inches):
    """Set the left indentation of a paragraph in inches

    Args:
        paragraph: The paragraph object to modify
        inches (float): Amount of indentation in inches

    Returns:
        paragraph: The modified paragraph object
    """
    paragraph.paragraph_format.left_indent_paragraph = Inches(inches)
    return paragraph


def add_bullet_list(document, ul_element, indentation=None):
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
        bullet_para.add_run(li.text)
        if indentation:
            left_indent_paragraph(bullet_para, indentation)
    return bullet_para


def add_formatted_paragraph(
    document,
    text,
    bold=False,
    italic=False,
    alignment=None,
    indentation=None,
    font_size=None,
):
    """Add a paragraph with consistent formatting

    Args:
        document: The Word document object
        text (str): Text content for the paragraph
        bold (bool, optional): Whether text should be bold. Defaults to False.
        italic (bool, optional): Whether text should be italic. Defaults to False.
        alignment: Paragraph alignment. Defaults to None.
        indentation (float, optional): Left indentation in inches. Defaults to None.
        font_size (int, optional): Font size in points. Defaults to None.

    Returns:
        paragraph: The created paragraph object
    """
    para = document.add_paragraph()
    run = para.add_run(text)

    # Apply formatting
    run.bold = bold
    run.italic = italic

    if alignment:
        para.alignment = alignment
    if indentation:
        left_indent_paragraph(para, indentation)
    if font_size:
        run.font.size = Pt(font_size)

    return para


def has_hr_before_section(section_h2):
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


def process_section(document, soup, section_type, page_break=False):
    """Process a standard section with heading and content

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        section_type: ResumeSection enum value
        page_break (bool, optional): Whether to add a page break before the section. Defaults to False.

    Returns:
        BeautifulSoup element or None: The section heading element if found, None otherwise
    """
    # Use case-insensitive comparison by using a lambda function
    section_h2 = soup.find("h2", string=lambda text: section_type.matches(text))
    if section_h2:
        # Add page break if requested
        if page_break:
            # Add a page break by inserting a run with a page break character
            p = document.add_paragraph()
            run = p.add_run()
            run.add_break(docx.enum.text.WD_BREAK.PAGE)

        # Add the section heading
        document.add_heading(
            section_type.docx_heading, level=MarkdownHeadingLevel.H2.value
        )
        return section_h2
    return None


def process_simple_section(document, soup, section_type, page_break=False, add_space=False):
    """Process sections with simple paragraph-based content like Education and Contact.
    These sections typically have paragraphs with some bold (strong) elements.

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        section_type: ResumeSection enum value
        page_break (bool, optional): Whether to add a page break before the section. Defaults to False.
        add_space (bool, optional): Whether to add a space paragraph after the section. Defaults to False.

    Returns:
        None
    """
    section_h2 = soup.find("h2", string=lambda text: section_type.matches(text))
    if not section_h2:
        return

    # Add page break if requested
    if page_break:
        # Add a page break by inserting a run with a page break character
        p = document.add_paragraph()
        run = p.add_run()
        run.add_break(docx.enum.text.WD_BREAK.PAGE)

    document.add_heading(section_type.docx_heading, level=MarkdownHeadingLevel.H2.value)
    current_element = section_h2.find_next_sibling()

    while current_element and current_element.name != "h2":
        if current_element.name == "p":
            para = document.add_paragraph()
            for child in current_element.children:
                if getattr(child, "name", None) == "strong":
                    para.add_run(f"{child.text}: ").bold = True
                else:
                    if child.string and child.string.strip():
                        para.add_run(child.string.strip())
        elif current_element.name == "ul":
            # Handle bullet lists if they appear
            add_bullet_list(document, current_element)

        current_element = current_element.find_next_sibling()

    # Add an extra space after the section if requested
    if add_space:
        add_space_paragraph(document)


def process_project_section(document, project_element, processed_elements):
    """Process a project/client section and its related elements

    Args:
        document: The Word document object
        project_element: BeautifulSoup element for the project heading
        processed_elements: Set of elements already processed

    Returns:
        set: Updated set of processed elements
    """
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

    project_para = document.add_paragraph()
    heading_level = MarkdownHeadingLevel.get_level_for_tag(project_element.name)
    if heading_level is not None:
        # Use heading style
        project_heading = document.add_heading("", level=heading_level)
        project_text = subsection.full_heading
        if project_info:
            project_text += " " + project_info
        project_heading.add_run(project_text)
    else:
        # Otherwise use paragraph with manual formatting
        project_text = subsection.full_heading
        if project_info:
            project_text += " " + project_info
        project_run = project_para.add_run(project_text)
        project_run.bold = subsection.bold
        project_run.italic = subsection.italic

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
            if heading_level is not None:
                # Use heading style
                resp_heading = document.add_heading(
                    h6_subsection.full_heading, level=heading_level
                )
                left_indent_paragraph(resp_heading, 0.25)  # Keep indentation for h6
            else:
                # Otherwise use paragraph with manual formatting
                resp_header = document.add_paragraph()
                left_indent_paragraph(resp_header, 0.25)  # Keep indentation for h6
                resp_run = resp_header.add_run(h6_subsection.full_heading)
                resp_run.bold = h6_subsection.bold
                resp_run.italic = h6_subsection.italic

            # Get the paragraph with responsibilities
            resp_element = next_element.find_next_sibling()
            if resp_element and resp_element.name == "p":
                resp_para = document.add_paragraph(resp_element.text)
                left_indent_paragraph(resp_para, 0.25)  # Keep indentation for content
                processed_elements.add(resp_element)

            processed_elements.add(next_element)

        # Additional Details
        elif h6_subsection == JobSubsection.ADDITIONAL_DETAILS:
            heading_level = MarkdownHeadingLevel.get_level_for_tag(next_element.name)
            if heading_level is not None:
                # Use heading style
                details_heading = document.add_heading(
                    h6_subsection.full_heading, level=heading_level
                )
                left_indent_paragraph(details_heading, 0.25)  # Keep indentation for h6
            else:
                # Otherwise use paragraph with manual formatting
                details_header = document.add_paragraph()
                left_indent_paragraph(details_header, 0.25)  # Keep indentation for h6
                details_run = details_header.add_run(h6_subsection.full_heading)
                details_run.bold = h6_subsection.bold
                details_run.italic = h6_subsection.italic
            processed_elements.add(next_element)

        # Bullet points
        elif next_element.name == "ul":
            for li in next_element.find_all("li"):
                bullet_para = document.add_paragraph(style="List Bullet")
                left_indent_paragraph(bullet_para, 0.5)  # Keep indentation for bullets
                bullet_para.add_run(li.text)
            processed_elements.add(next_element)

        next_element = next_element.find_next_sibling()

    return processed_elements


def process_job_entry(document, job_element, processed_elements):
    """Process a job entry (h3) and its related elements

    Args:
        document: The Word document object
        job_element: BeautifulSoup element for the job heading
        processed_elements: Set of elements already processed

    Returns:
        set: Updated set of processed elements
    """
    job_title = job_element.text.strip()

    # Use proper heading style instead of manual paragraph formatting
    heading_level = MarkdownHeadingLevel.get_level_for_tag(job_element.name)
    document.add_heading(job_title, level=heading_level)

    # Find company name (h4) if it exists
    next_element = job_element.find_next_sibling()

    # Add company if it exists
    if next_element and next_element.name == "h4":
        company_name = next_element.text.strip()

        # Use proper heading style for company name
        company_heading_level = MarkdownHeadingLevel.get_level_for_tag(
            next_element.name
        )
        document.add_heading(company_name, level=company_heading_level)

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


def process_certifications(document, soup, section_type, page_break=False):
    """Process the certifications section with its specific structure

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        section_type: ResumeSection enum value (should be CERTIFICATIONS)
        page_break (bool, optional): Whether to add a page break before the section. Defaults to False.

    Returns:
        None
    """
    section_h2 = soup.find("h2", string=lambda text: section_type.matches(text))
    if not section_h2:
        return

    # Add page break if requested
    if page_break:
        p = document.add_paragraph()
        run = p.add_run()
        run.add_break(docx.enum.text.WD_BREAK.PAGE)

    document.add_heading(section_type.docx_heading, level=MarkdownHeadingLevel.H2.value)
    current_element = section_h2.find_next_sibling()

    while current_element and current_element.name != "h2":
        # Process certification name (h3)
        if current_element.name == "h3":
            cert_name = current_element.text.strip()
            cert_heading = document.add_heading(cert_name, level=MarkdownHeadingLevel.H3.value)

            # Look for next elements - either blockquote or organization info directly
            next_element = current_element.find_next_sibling()

            # Handle blockquote (optional)
            if next_element and next_element.name == "blockquote":
                # Process the blockquote contents
                process_certification_blockquote(document, next_element)

            # If no blockquote, look for organization info directly
            elif next_element:
                # Try to find organization info (could be bold text or heading)
                if next_element.name in ["h4", "h5", "h6", "p"] and next_element.find("strong"):
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
                            add_hyperlink(date_para, date_text, parent_a["href"])
                        else:
                            date_run = date_para.add_run(date_text)
                            date_run.italic = True

            # Add spacing after each certification
            add_space_paragraph(document)

        current_element = current_element.find_next_sibling()


def process_certification_blockquote(document, blockquote):
    """Process the contents of a certification blockquote

    Args:
        document: The Word document object
        blockquote: BeautifulSoup blockquote element containing certification details

    Returns:
        None
    """
    # Process the organization name (first element)
    org_element = blockquote.find(['h4', 'h5', 'h6', 'strong'])
    if org_element:
        if org_element.name.startswith('h'):
            # It's a heading
            heading_level = MarkdownHeadingLevel.get_level_for_tag(org_element.name)
            document.add_heading(org_element.text.strip(), level=heading_level)
        elif org_element.name == 'strong':
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
        if hasattr(item, 'name') and item.name and item.name.startswith('h') and item != org_element:
            heading_level = MarkdownHeadingLevel.get_level_for_tag(item.name)
            document.add_heading(item.text.strip(), level=heading_level)
            continue

        # Check if it's the date with italics (should be last non-empty element)
        if hasattr(item, 'find') and item.find('em'):
            em_tag = item.find('em')
            date_text = em_tag.text.strip()
            date_para = document.add_paragraph()

            # Check if inside a hyperlink
            parent_a = em_tag.find_parent('a')
            if parent_a and parent_a.get('href'):
                add_hyperlink(date_para, date_text, parent_a['href'])
            else:
                date_run = date_para.add_run(date_text)
                date_run.italic = True
            continue

        # Handle normal text content (like "Some details")
        if isinstance(item, str) and item.strip():
            # It's a direct text node
            document.add_paragraph(item.strip())
        elif hasattr(item, 'name'):
            if item.name == 'p':
                # It's a paragraph
                document.add_paragraph(item.text.strip())
            elif item.name == 'ul':
                # It's a bullet list
                add_bullet_list(document, item)


def add_hyperlink(paragraph, text, url):
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
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the hyperlink element
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id)

    # Create a run inside the hyperlink
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    # Add text style for hyperlinks (blue and underlined)
    color = docx.oxml.shared.OxmlElement('w:color')
    color.set(docx.oxml.shared.qn('w:val'), '0000FF')
    rPr.append(color)

    u = docx.oxml.shared.OxmlElement('w:u')
    u.set(docx.oxml.shared.qn('w:val'), 'single')
    rPr.append(u)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    # Add the hyperlink to the paragraph
    paragraph._p.append(hyperlink)

    return hyperlink


def add_horizontal_line_simple(document):
    """Add a simple horizontal line to the document using underscores

    Args:
        document: The Word document object

    Returns:
        None
    """
    p = document.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.add_run("_" * 50).bold = True

    # Add some space after the line
    # add_space_paragraph(document)


def add_space_paragraph(document):
    """Add a paragraph with extra space after it

    Args:
        document: The Word document object

    Returns:
        None
    """
    p = document.add_paragraph()
    p.paragraph_format.space_after = Pt(12)


if __name__ == "__main__":
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
      python resume_md_to_docx.py -i resume.md
          - Converts resume.md to "My ATS Resume.docx"

      python resume_md_to_docx.py -i resume.md -o resume.docx
          - Converts resume.md to resume.docx

      python resume_md_to_docx.py --input ~/Documents/resume.md -o resume.docx
          - Uses a resume from a different location
    """

    # Parse command line arguments with enhanced help
    parser = argparse.ArgumentParser(
        description=program_description,
        epilog=epilog_text,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "-i", "--input", dest="input_file", help="Input markdown file (required)"
    )
    parser.add_argument(
        "-o",
        "--output",
        dest="output_file",
        default="My ATS Resume.docx",
        help='Output Word document (default: "My ATS Resume.docx")',
    )

    args = parser.parse_args()

    # Show welcome screen if requested
    if not args.input_file:
        # display the help message
        parser.print_help()
    else:
        # Use the parsed arguments
        result = create_ats_resume(args.input_file, args.output_file)
        print(f"🎉 ATS-friendly resume created: {result} 🎉")
