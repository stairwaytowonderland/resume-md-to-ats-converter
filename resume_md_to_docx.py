# pip install python-docx markdown beautifulsoup4
# python resume_md_to_docx.py

import argparse
from enum import Enum

import markdown
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt


class ResumeSection(Enum):
    """Maps markdown heading titles to their corresponding document headings"""

    ABOUT = ("About", "PROFESSIONAL SUMMARY")
    SKILLS = ("Top Skills", "CORE SKILLS")
    EXPERIENCE = ("Experience", "PROFESSIONAL EXPERIENCE")
    EDUCATION = ("Education", "EDUCATION")
    CONTACT = ("Contact", "CONTACT INFORMATION")

    def __init__(self, markdown_heading, docx_heading):
        self.markdown_heading = markdown_heading
        self.docx_heading = docx_heading

class JobSubsection(Enum):
    """Maps markdown subsection headings to their corresponding document headings"""

    KEY_SKILLS = ("Key Skills", "Technical Skills: ")
    SUMMARY = ("Summary", "Summary:")
    INTERNAL = ("Internal", "Internal")
    PROJECT_CLIENT = ("Project/Client", "Project/Client")
    RESPONSIBILITIES = ("Responsibilities Overview", "Responsibilities:")
    ADDITIONAL_DETAILS = ("Additional Details", "Additional Details:")

    def __init__(self, markdown_heading, docx_heading):
        self.markdown_heading = markdown_heading
        self.docx_heading = docx_heading

def create_ats_resume(md_file, output_file):
    """Convert markdown resume to ATS-friendly Word document"""
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
    title = document.add_heading(name, 0)
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
    about_h2 = process_section(document, soup, ResumeSection.ABOUT)
    if about_h2:
        current_p = about_h2.find_next_sibling()

        # Add all paragraphs until next h2
        while current_p and current_p.name != "h2":
            if current_p.name == "p" and not current_p.find(
                "strong", string="Some highlights"
            ):
                document.add_paragraph(current_p.text)
            elif current_p.name == "p" and current_p.find(
                "strong", string="Some highlights"
            ):
                # Add the "Some highlights" as a separate paragraph
                highlights_run = add_formatted_paragraph(
                    document, "Some highlights", bold=True
                ).add_run()
            elif current_p.name == "ul":
                add_bullet_list(document, current_p)
            current_p = current_p.find_next_sibling()

    # Process Top Skills section
    skills_h2 = process_section(document, soup, ResumeSection.SKILLS)
    if skills_h2:
        current_p = skills_h2.find_next_sibling()
        if current_p and current_p.name == "p":
            skills = [s.strip() for s in current_p.text.split("•")]
            skills = [s for s in skills if s]
            add_formatted_paragraph(document, " | ".join(skills))

    # Process Experience section
    experience_h2 = soup.find("h2", string=ResumeSection.EXPERIENCE.markdown_heading)
    if experience_h2:
        document.add_heading(ResumeSection.EXPERIENCE.docx_heading, level=1)

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
                    # This is a position title
                    position_para = document.add_paragraph()
                    position_run = position_para.add_run(position_title)
                    position_run.bold = True
                    position_run.font.size = Pt(12)  # Consistent size for all h4s

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

            elif (
                current_element.name == "h5" and
                JobSubsection.PROJECT_CLIENT.markdown_heading in current_element.text
            ):
                processed_elements = process_project_section(
                    document, current_element, processed_elements
                )

            elif current_element.name == "h5" and JobSubsection.SUMMARY.markdown_heading in current_element.text:
                summary_para = document.add_paragraph()
                summary_run = summary_para.add_run(JobSubsection.SUMMARY.docx_heading)
                summary_run.bold = True

                # Find the next element with summary content
                next_element = current_element.find_next_sibling()
                if next_element and next_element.name == "p":
                    document.add_paragraph(next_element.text)
                    processed_elements.add(next_element)

            elif current_element.name == "h5" and JobSubsection.INTERNAL.markdown_heading in current_element.text:
                internal_para = document.add_paragraph()
                internal_run = internal_para.add_run(JobSubsection.INTERNAL.docx_heading)
                internal_run.bold = True

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
                            bullet_para = document.add_paragraph(style="List Bullet")
                            left_indent_paragraph(
                                bullet_para, 0.5
                            )  # Keep indentation for bullets
                            bullet_para.add_run(li.text)
                        processed_elements.add(next_element)

                    next_element = next_element.find_next_sibling()

            elif current_element.name == "h5" and JobSubsection.KEY_SKILLS.markdown_heading in current_element.text:
                skills_para = document.add_paragraph()
                skills_run = skills_para.add_run(JobSubsection.KEY_SKILLS.docx_heading)
                skills_run.bold = True

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

            # Standalone Responsibilities Overview (not under Project/Client)
            elif (
                current_element.name == "h6" and
                JobSubsection.RESPONSIBILITIES.markdown_heading in current_element.text and current_element not in processed_elements
            ):
                resp_header = document.add_paragraph()
                resp_run = resp_header.add_run(JobSubsection.RESPONSIBILITIES.docx_heading)
                resp_run.bold = True

                # Get content
                next_element = current_element.find_next_sibling()
                if next_element and next_element.name == "p":
                    document.add_paragraph(next_element.text)
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
    process_simple_section(document, soup, ResumeSection.EDUCATION)

    # Add an extra space after the last section
    add_space_paragraph(document)

    # Add contact information
    process_simple_section(document, soup, ResumeSection.CONTACT)

    # Save the document
    document.save(output_file)
    return output_file


def left_indent_paragraph(paragraph, inches):
    """Set the left indentation of a paragraph in inches"""
    paragraph.paragraph_format.left_indent_paragraph = Inches(inches)
    return paragraph


def add_bullet_list(document, ul_element, indentation=None):
    """Add bullet points from an unordered list element"""
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
    """Add a paragraph with consistent formatting"""
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


def process_section(document, soup, section_type):
    """Process a standard section with heading and content

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        section_type: ResumeSection enum value
    """
    section_h2 = soup.find("h2", string=section_type.markdown_heading)
    if section_h2:
        document.add_heading(section_type.docx_heading, level=1)
        return section_h2
    return None


def process_simple_section(document, soup, section_type):
    """
    Process sections with simple paragraph-based content like Education and Contact.
    These sections typically have paragraphs with some bold (strong) elements.

    Args:
        document: The Word document object
        soup: BeautifulSoup object of the HTML content
        section_type: ResumeSection enum value
    """
    section_h2 = soup.find("h2", string=section_type.markdown_heading)
    if not section_h2:
        return

    document.add_heading(section_type.docx_heading, level=1)
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


def process_project_section(document, project_element, processed_elements):
    """Process a project/client section and its related elements"""
    # Get the next element to see if it contains the project details
    next_element = project_element.find_next_sibling()
    project_info = ""

    # If next element is a paragraph, it might contain the project name/duration
    if next_element and next_element.name == "p":
        project_info = next_element.text.strip()
        processed_elements.add(next_element)
        next_element = next_element.find_next_sibling()

    project_para = document.add_paragraph()
    project_text = JobSubsection.PROJECT_CLIENT.docx_heading
    if project_info:
        project_text += f": {project_info}"
    project_run = project_para.add_run(project_text)
    project_run.bold = True
    project_run.italic = True

    # Process next elements under this project until another section
    while (
        next_element
        and next_element.name not in ["h2", "h3", "h4"]
        and not (next_element.name == "h5")
    ):
        # Responsibilities Overview
        if (
            next_element.name == "h6"
            and "Responsibilities Overview" in next_element.text
        ):
            resp_header = document.add_paragraph()
            left_indent_paragraph(resp_header, 0.25)  # Keep indentation for h6
            resp_run = resp_header.add_run("Responsibilities:")
            resp_run.bold = True

            # Get the paragraph with responsibilities
            resp_element = next_element.find_next_sibling()
            if resp_element and resp_element.name == "p":
                resp_para = document.add_paragraph(resp_element.text)
                left_indent_paragraph(resp_para, 0.25)  # Keep indentation for content
                processed_elements.add(resp_element)

            processed_elements.add(next_element)

        # Additional Details
        elif next_element.name == "h6" and JobSubsection.ADDITIONAL_DETAILS.markdown_heading in next_element.text:
            details_header = document.add_paragraph()
            left_indent_paragraph(details_header, 0.25)  # Keep indentation for h6
            details_run = details_header.add_run(JobSubsection.ADDITIONAL_DETAILS.docx_heading)
            details_run.bold = True
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
    """Process a job entry (h3) and its related elements"""
    job_title = job_element.text.strip()

    # Add job title with consistent formatting
    job_para = document.add_paragraph()
    job_run = job_para.add_run(job_title)
    job_run.bold = True
    job_run.font.size = Pt(14)  # Consistent size for all h3s

    # Find company name (h4) if it exists
    next_element = job_element.find_next_sibling()

    # Add company if it exists
    if next_element and next_element.name == "h4":
        company_name = next_element.text.strip()

        # Add company with consistent formatting
        company_para = document.add_paragraph()
        company_run = company_para.add_run(company_name)
        company_run.bold = True

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


def add_horizontal_line_simple(document):
    """Add a simple horizontal line to the document using underscores"""
    p = document.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.add_run("_" * 50).bold = True

    # Add some space after the line
    # add_space_paragraph(document)


def add_space_paragraph(document):
    """Add a simple horizontal line to the document using underscores"""
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
