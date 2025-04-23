import argparse
import json
import logging
import os
import tempfile
from pathlib import Path
from typing import Any

import yaml
from flask import Flask, send_from_directory
from flask.wrappers import Response
from flask_restx import Api, Resource, fields
from werkzeug.datastructures import FileStorage

# Import functionality from your script
from resume_md_to_docx import *

logging.basicConfig(
    level=logging.INFO,
    datefmt="%Y-%m-%d %H:%M:%S",
    format="%(asctime)s.%(msecs)d %(levelname)-8s [%(processName)s] [%(threadName)s] %(filename)s:%(funcName)s:%(lineno)d --- %(message)s",
)

SCRIPT_DIR = Path(__file__).parent.parent


class ApiConfig:
    """Application configuration class"""

    API_CONFIG_FILE = SCRIPT_DIR / "api_config.yaml"

    def __init__(self):
        """Initialize the application configuration"""

        self._file = ApiConfig.API_CONFIG_FILE.absolute().resolve()
        self._config = self.load_app_config()

        # Required settings
        self._server = self._config.get("server")

    @property
    def file(self) -> Path:
        """Get document default settings

        Returns:
            dict: Document defaults configuration
        """
        return Path(self._file)

    @property
    def config(self) -> dict:
        """Get the entire configuration dictionary

        Returns:
            dict: Complete configuration dictionary
        """
        return self._config

    @property
    def server(self) -> str:
        """Get the server name for the API

        Returns:
            str: Server name for the API
        """
        return self._server

    @property
    def mimetypes(self) -> dict:
        """Get mimetypes settings

        Returns:
            dict: Mimetypes configuration
        """
        return self._config.get("mimetypes", {})

    @property
    def cors(self) -> dict:
        """Get cors settings

        Returns:
            dict: Cors configuration
        """
        return self._config.get("cors", {})

    @property
    def logging(self) -> dict:
        """Get logging settings

        Returns:
            dict: Logging configuration
        """
        return self._config.get("logging", {})

    @property
    def output(self) -> dict:
        """Get output settings

        Returns:
            dict: Output configuration
        """
        return self._config.get("output", {})

    # Load API configuration
    def load_app_config(self) -> dict[str, Any]:
        """Load API configuration from app_config.yaml

        Returns:
            dict: Application configuration
        """
        if os.path.exists(self._file):
            try:
                with open(self._file, "r", encoding="utf-8", errors="replace") as f:
                    return yaml.safe_load(f)
            except Exception as e:
                print(f"Error loading app config: {e}")
                return {}
        else:
            print(f"Warning: {self._file} not found, using defaults")
            return {}


class BaseApp:
    """Base class for Flask application"""

    def configure_logging(app: Flask, api_config: ApiConfig) -> None:
        # Configure logging
        log_level_name = api_config.logging.get("level", "INFO")
        print(f"Log level: {log_level_name}")
        app.logger.setLevel(getattr(logging, log_level_name))

    def __init__(self, app: Flask, app_config: ApiConfig):
        """Initialize the API

        Args:
            app (Flask): Flask application instance
            app_config (ApiConfig): Application configuration instance
        """
        self.app = app
        self.app_config = app_config
        self.api = Api(
            app,
            version="1.0",
            title="Resume Markdown to DOCX API",
            description="API for converting markdown resumes to ATS-friendly formats",
            doc="/swagger",
        )
        self.ns = self.api.namespace(
            "convert", description="Resume conversion operations"
        )

        App.configure_logging(app, app_config)
        App.configure_cors(app, app_config)


class App(BaseApp):
    """API class for handling resume conversion"""

    def configure_cors(app: Flask, api_config: ApiConfig) -> None:
        # Configure CORS if enabled
        cors_config = api_config.cors
        if api_config.cors.get("enabled", False):
            from flask_cors import CORS

            app.logger.info(f"Configuring CORS with: {cors_config}")
            CORS(
                app,
                resources={
                    r"/convert/*": {
                        "origins": cors_config.get("origins", "*"),
                        "expose_headers": cors_config.get(
                            "expose_headers", ["Content-Disposition"]
                        ),
                    }
                },
                supports_credentials=cors_config.get("supports_credentials", "*"),
            )
        else:
            app.logger.info("CORS disabled")

    def __init__(self, app: Flask, app_config: ApiConfig):
        """Initialize the API

        Args:
            app (Flask): Flask application instance
            app_config (ApiConfig): Application configuration instance
        """
        super().__init__(app, app_config)

        self.error_response_model = self.api.model(
            "Response",
            {
                "success": fields.Boolean(
                    description="Whether the operation was successful"
                ),
                "message": fields.String(description="Status message"),
            },
        )
        self.arg_parser = self.api.parser()
        self.arg_parser.add_argument(
            "input_file",
            location="files",
            type=FileStorage,
            required=True,
            help="Markdown resume file",
        )
        self.arg_parser.add_argument(
            "paragraph_headings",
            action="append",
            choices=["h3", "h4", "h5", "h6"],
            help="Heading levels to render as paragraphs",
        )
        self.arg_parser.add_argument(
            "config_options",
            type=str,
            required=False,
            help="JSON string with configuration overrides",
        )

    def _check_extension(self, expected_extension: str, filename: Path = None) -> bool:
        # Check if the file has a valid extension
        if filename and not filename.suffix == f".{expected_extension}":
            raise ValueError(
                f"Invalid file extension: .{expected_extension} is expected"
            )
        return True

    def _error_response(
        self, code: int, error: object, message: str = None
    ) -> tuple[dict[str, Any], int]:
        """Return a JSON error response

        Args:
            message (str): Error message to return
            level (int): HTTP status code

        Returns:
            tuple: JSON response with error message and status code
        """
        msg = f"{message}: {str(error)}" if message else str(error)
        self.app.logger.error(msg)
        return {
            "success": False,
            "message": msg,
        }, code

    def _response(
        self,
        md_input_path: Path,
        docx_output_path: Path,
        paragraph_headings: list[str],
        output_formats: list[str],
        config_loader: ConfigLoader,
        filename: Path = None,
    ) -> Response:
        self.app.logger.info(f"Markdown input file: {md_input_path}")
        self.app.logger.info(f"Docx output file: {docx_output_path}")

        # Create paragraph_style_headings dictionary
        paragraph_style_headings = {}
        for heading_level in paragraph_headings:
            if heading_level in ["h3", "h4", "h5", "h6"]:
                paragraph_style_headings[heading_level] = True

        try:
            # Convert markdown to DOCX
            docx_path = create_ats_resume(
                md_input_path,
                docx_output_path,
                config_loader=config_loader,
                paragraph_style_headings=paragraph_style_headings,
            )

            # Track created files and file to return
            output_file = None
            mimetype = None

            # Process DOCX if requested
            if DOCX_EXTENSION in output_formats:
                if self._check_extension(DOCX_EXTENSION, filename):
                    self.app.logger.info(f"Output extension: {DOCX_EXTENSION}")
                    if os.path.exists(docx_path):
                        output_file = docx_path
                        mimetype = DOCX_MIMETYPE

            # Process PDF if requested (convert from the generated DOCX)
            elif PDF_EXTENSION in output_formats:
                if self._check_extension(PDF_EXTENSION, filename):
                    self.app.logger.info(f"Output extension: {PDF_EXTENSION}")
                    pdf_path = convert_to_pdf(docx_path)
                    if pdf_path and os.path.exists(pdf_path):
                        self.app.logger.info(f"PDF conversion successful: {pdf_path}")
                        output_file = pdf_path
                        mimetype = PDF_MIMETYPE

            else:
                raise ValueError("Invalid output format specified")

            # If we don't have a file to return, that's an error
            self.app.logger.info(f"Output file: {output_file}")
            if not output_file or not os.path.exists(output_file):
                raise Exception(f"Failed to generate output file: {output_file}")

            # Return the appropriate file directly from the temp directory
            # Add explicit filename in Content-Disposition header for curl -O
            download_name = os.path.basename(output_file)
            self.app.logger.info(f"Successfully created: {output_file}")

            # response = send_file(
            #     output_file,
            #     as_attachment=True,
            #     download_name=download_name,
            #     mimetype=mimetype,
            # )

            # More secure than send_file
            response = send_from_directory(
                directory=docx_output_path.parent,
                path=output_file.name,
                as_attachment=True,
                download_name=download_name,
                mimetype=mimetype,
            )

            # Force proper filename in Content-Disposition header
            # This is more explicit than Flask's internal handling
            response.headers["Content-Disposition"] = (
                f'attachment; filename="{download_name}"'
            )

            # Add additional headers to help browsers handle the download properly
            # response.headers["X-Content-Type-Options"] = "nosniff"

            return response

        except ValueError as e:
            return self._error_response(400, f"Value error: {str(e)}")
        except FileNotFoundError as e:
            self._error_response(404, e, "File not found")
        except Exception as e:
            return self._error_response(400, f"Error: {str(e)}")

    def post(
        self,
        output_format: str = DEFAULT_OUTPUT_FORMAT,
        filename: str = None,
    ) -> Response | tuple[dict[str, Any], int]:
        """Convert markdown resume to DOCX and optionally PDF

        Args:
            output_format (str): Output format to generate (docx or pdf)
            use_output_dir (bool): Whether to use the output directory for saving files

        Returns:
            Response: Flask response with the generated file
        """
        if filename:
            filename = Path(filename)

        # Always load the config
        config_loader = ConfigLoader()

        # Get the uploaded file and parameters
        args = api.arg_parser.parse_args()
        input_file = args["input_file"]
        output_formats = (
            [output_format] if isinstance(output_format, str) else output_format
        )
        output_name = Path(input_file.filename).stem
        paragraph_headings = args["paragraph_headings"] or []

        # Parse config_options if provided
        config_data = {}
        if args["config_options"]:
            try:
                config_data = json.loads(args["config_options"])
            except json.JSONDecodeError as e:
                return self._error_response(
                    400, e, "Invalid JSON in config_options parameter"
                )

        _resolve_config_helper(self.app, config_loader, config_data)
        self.app.logger.debug(f"Configuration loaded: {config_loader.config}")
        temp_dir_enabled = api_config.output.get("use_temp_directory", True)
        self.app.logger.info(f"Temporary directory enabled: {temp_dir_enabled}")

        base_input_filename = input_file.filename
        base_output_filename = f"{output_name}.{DOCX_EXTENSION}"

        if temp_dir_enabled:
            with tempfile.TemporaryDirectory() as temp_dir:

                # Save the uploaded file
                temp_input_path = Path(temp_dir) / base_input_filename
                input_file.save(temp_input_path)

                # Prepare output paths directly in the temporary directory
                temp_output_path = Path(temp_dir) / base_output_filename

                return self._response(
                    temp_input_path,
                    temp_output_path,
                    paragraph_headings,
                    output_formats,
                    config_loader,
                    filename,
                )
        else:
            output_path = DEFAULT_OUTPUT_DIR / (filename or base_output_filename)
            return self._response(
                base_input_filename,
                output_path,
                paragraph_headings,
                output_formats,
                config_loader,
                filename,
            )


# Load application configuration
api_config = ApiConfig()

# Extract configuration values with defaults
DOCX_MIMETYPE = api_config.mimetypes.get(
    "docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
PDF_MIMETYPE = api_config.mimetypes.get("mimetypes", {}).get("pdf", "application/pdf")

# Extract configuration values with defaults
SERVER_NAME = f"{api_config.server.get('host')}:{api_config.server.get('port')}"

app = Flask(__name__.split(".")[0])
api = App(app, api_config)


def _resolve_config_helper(
    app: Flask,
    config_loader: ConfigLoader,
    config_options: dict[str, Any] = None,
) -> ConfigLoader:
    """Merge the provided config options with the existing config

    Args:
        app (Flask): Flask application instance
        config_loader (ConfigLoader): Existing config loader
        config_options (dict): Configuration options to merge
    """
    if config_options:
        app.logger.info(f"Merging custom configuration: {config_options}")

        # Update top-level config sections
        for section_key, section_values in config_options.items():
            if section_key in config_loader.config:
                # If section exists in default config, update it
                if isinstance(section_values, dict) and isinstance(
                    config_loader.config[section_key], dict
                ):
                    app.logger.debug(
                        f"Merging section '{section_key}' with values: {section_values}"
                    )
                    config_loader.config[section_key].update(section_values)
                else:
                    # Replace the entire section if it's not a mergeable dictionary
                    app.logger.debug(
                        f"Replacing section '{section_key}' with values: {section_values}"
                    )
                    config_loader.config[section_key] = section_values
            else:
                # Add new section if it doesn't exist
                app.logger.debug(
                    f"Adding new section '{section_key}' with values: {section_values}"
                )
                config_loader.config[section_key] = section_values


@api.ns.route("/docx")
@api.ns.route(f"/docx/<path:name>")
class ConvertDocxResource(Resource):
    @api.ns.doc("convert_markdown")
    @api.ns.expect(api.arg_parser)
    @api.ns.response(200, "Success - Returns DOCX file download")
    @api.ns.response(400, "Bad Request", api.error_response_model)
    @api.ns.response(404, "File Not Found", api.error_response_model)
    @api.ns.response(500, "Server Error", api.error_response_model)
    def post(self, name: str = None) -> Response:
        """Convert markdown resume to DOCX

        Args:
            name (str): Optional filename for the output DOCX

        Returns:
            Response: Flask response with the generated DOCX file
        """
        return api.post(DOCX_EXTENSION, name)


@api.ns.route("/pdf")
@api.ns.route(f"/pdf/<path:name>")
class ConvertPdfResource(Resource):
    @api.ns.doc("convert_markdown")
    @api.ns.expect(api.arg_parser)
    @api.ns.response(200, "Success - Returns PDF file download")
    @api.ns.response(400, "Bad Request", api.error_response_model)
    @api.ns.response(404, "File Not Found", api.error_response_model)
    @api.ns.response(500, "Server Error", api.error_response_model)
    def post(self, name: str = None) -> None:
        """Convert markdown resume to PDF

        Args:
            name (str): Optional filename for the output PDF

        Returns:
            Response: Flask response with the generated PDF file
        """
        return api.post(PDF_EXTENSION, name)


if __name__ == "__main__":
    # Program description and epilog
    program_description = """
Resume Markdown to DOCX API
--------------------------------
This API converts markdown resumes to ATS-friendly formats (DOCX and PDF).
It provides endpoints for converting markdown files to DOCX and PDF formats.
"""

    epilog_text = """
Example usage:
  # Start the API server
  python api.py --config api_config.yaml --debug

  # Convert a markdown resume to PDF
  curl -X POST "http://localhost:3000/convert/pdf" \\
    -H "Content-Type: multipart/form-data" \\
    -F "input_file=@resume.md" \\
    -F "paragraph_headings=h5" \\
    -F "paragraph_headings=h6" \\
    -F "config_options={\"document_styles\": {\"Subtitle\": {\"font_name\": \"Helvetica Neue\"}}}"
"""

    # Parse command line arguments with enhanced help
    parser = argparse.ArgumentParser(
        description=program_description,
        epilog=epilog_text,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "-c",
        "--config",
        dest="config_file",
        help="Path to YAML configuration file",
        default=ApiConfig.API_CONFIG_FILE,
    )

    parser.add_argument(
        "--debug",
        action="store_true",
        dest="debug",
        help="Enable debug mode for the Flask application",
        default=False,
    )

    args = parser.parse_args()

    app.debug = args.debug
    app.config["SERVER_NAME"] = SERVER_NAME
    app.run(port=api_config.server.get("port"))
