import argparse
import json
import logging
import os
import tempfile
import uuid
from pathlib import Path
from typing import Any

import yaml
from flask import Flask, request, send_from_directory
from flask.wrappers import Response
from flask_restx import Api, Resource, fields
from werkzeug.datastructures import FileStorage

from src.resume_md_to_docx import *

# To be able to run as `python src/api.py` (or `python3 api.py`):
# if __name__ == "__main__":
#     from resume_md_to_docx import *
# else:
#     from src.resume_md_to_docx import *

logging.basicConfig(
    level=logging.INFO,
    datefmt="%Y-%m-%d %H:%M:%S",
    format="%(asctime)s.%(msecs)d %(levelname)-8s [%(processName)s] [%(threadName)s] %(filename)s:%(funcName)s:%(lineno)d --- %(message)s",
)

SCRIPT_DIR = Path(__file__).parent
API_CONFIG_FILE = Path("api_config.yaml")


class ApiConfig:
    """Application configuration class"""

    def __init__(self, api_config_file: Path):
        """Initialize the application configuration

        Args:
            api_config_file (Path): Path to the API configuration file
        """

        self._config_file = api_config_file
        self._config_file_realpath = api_config_file.absolute().resolve()
        self._config = self.load_app_config()

        # Required settings
        self._server = self._config.get("server")

    @property
    def config_file(self) -> Path:
        """Get document default settings

        Returns:
            dict: Document defaults configuration
        """
        return Path(self._config_file)

    @property
    def config_file_realpath(self) -> Path:
        """Get document default settings

        Returns:
            dict: Document defaults configuration
        """
        return Path(self._config_file_realpath)

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
    def mimetypes(self) -> dict[str, list[str]]:
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
    def input(self) -> dict:
        """Get input settings

        Returns:
            dict: Input configuration
        """
        return self._config.get("input", {})

    @property
    def output(self) -> dict:
        """Get output settings

        Returns:
            dict: Output configuration
        """
        return self._config.get("output", {})

    # Load API configuration
    def load_app_config(self) -> dict[str, Any]:
        """Load API configuration from api_config.yaml

        Returns:
            dict: Application configuration
        """
        if os.path.exists(self._config_file_realpath):
            try:
                with open(
                    self._config_file_realpath, "r", encoding="utf-8", errors="replace"
                ) as f:
                    return yaml.safe_load(f)
            except Exception as e:
                print(f"Error loading app config: {e}")
                return {}
        else:
            print(f"Warning: {self._config_file_realpath} not found, using defaults")
            return {}


class BaseApi:
    """Base class for Flask application"""

    def __init__(self, api_config_file: Path):
        """Initialize the API Base

        Args:
            api_config_file (Path): Path to the API configuration file
        """

        # Load application configuration
        api_config = ApiConfig(api_config_file)

        # Create Flask application
        app = Flask(__name__.split(".")[0])

        self._app = app
        self._api_config = api_config
        self._api = Api(
            app,
            version="1.0",
            title="Resume Markdown to DOCX API",
            description="API for converting markdown resumes to ATS-friendly formats",
            doc="/swagger",
        )
        self._ns = self._api.namespace(
            "convert", description="Resume conversion operations"
        )

        self._host = self._api_config.server.get("host")
        self._port = self._api_config.server.get("port")
        self._app.config["SERVER_NAME"] = f"{self._host}:{self._port}"

        self._app.logger.debug(f"API host: {self._host}")
        self._app.logger.debug(f"API port: {self._port}")
        self._app.logger.debug(f"API mimetypes: {self._api_config.mimetypes}")
        self._app.logger.debug(f"API cors: {self._api_config.cors}")
        self._app.logger.debug(f"API output: {self._api_config.output}")

        self._arg_parser = self._api.parser()

        self._configure_logging()
        self._configure_cors()

    @property
    def app(self) -> Flask:
        """Get the Flask application instance

        Returns:
            Flask: Flask application instance
        """
        return self._app

    @property
    def api(self) -> Api:
        """Get the API instance

        Returns:
            Api: API instance
        """
        return self._api

    @property
    def api_config(self) -> ApiConfig:
        """Get the API configuration instance

        Returns:
            ApiConfig: API configuration instance
        """
        return self._api_config

    @property
    def ns(self) -> Api:
        """Get the namespace instance

        Returns:
            Api: Namespace instance
        """
        return self._ns

    @property
    def arg_parser(self) -> argparse.ArgumentParser:
        """Get the argument parser instance

        Returns:
            argparse.ArgumentParser: Argument parser instance
        """
        return self._arg_parser

    @property
    def host(self) -> str:
        """Get the host for the API

        Returns:
            str: Host for the API
        """
        return self._host

    @property
    def port(self) -> int:
        """Get the port for the API

        Returns:
            int: Port for the API
        """
        return self._port

    def run(self, program_description: str = None, epilog_text: str = None) -> None:
        """Run the Flask application

        Args:
            program_description (str): Description of the program
            epilog_text (str): Epilog text for the help message
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
            default=self._api_config.config_file,
        )

        parser.add_argument(
            "--debug",
            action="store_true",
            dest="debug",
            help="Enable debug mode for the Flask application",
            default=False,
        )

        args = parser.parse_args()

        self._app.debug = args.debug
        self._app.run()

    def _configure_logging(self) -> None:
        """Configure logging for the API"""
        log_level_name = self._api_config.logging.get("level", "INFO")
        self._app.logger.setLevel(getattr(logging, log_level_name))
        self._app.logger.info(f"Logging level set to {log_level_name}")

    def _configure_cors(self) -> None:
        """Configure CORS for the API"""
        # Configure CORS if enabled
        cors_config = self._api_config.cors
        if self._api_config.cors.get("enabled", False):
            from flask_cors import CORS

            self._app.logger.info(f"Configuring CORS with: {cors_config}")
            CORS(
                self._app,
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
            self._app.logger.info("CORS disabled")


class App(BaseApi):
    """API class for handling resume conversion"""

    TEXT_SCHEMA = {
        "type": "string",
        "format": "text",
        "description": "Raw markdown content for conversion",
        "nullable": True,  # This makes the field optional,
    }

    def __init__(self, api_config_file: Path):
        """Initialize the API

        Args:
            app (Flask): Flask application instance
            api_config (ApiConfig): Application configuration instance
        """
        super().__init__(api_config_file)

        self._arg_parser.add_argument(
            "input_file",
            location="files",
            type=FileStorage,
            required=False,
            help="Markdown resume file",
        )
        self._arg_parser.add_argument(
            "config_options",
            type=str,
            required=False,
            help="JSON string with configuration overrides",
        )

    @property
    def response_model(self) -> dict:
        """Get the standard response model

        Returns:
            dict: Response model
        """
        return self._api.model(
            "Response",
            {
                "success": fields.Boolean(
                    description="Whether the operation was successful"
                ),
                "message": fields.String(description="Status message"),
            },
        )

    def _check_extension(self, expected_extension: str, filename: Path = None) -> bool:
        """Check if the file has a valid extension

        Args:
            expected_extension (str): Expected file extension
            filename (Path): File name to check

        Returns:
            bool: True if the file has a valid extension, False otherwise

        Raises:
            ValueError: If the file extension is invalid
        """
        # Check if the file has a valid extension
        if filename and not filename.suffix == f".{expected_extension}":
            raise ValueError(
                f"Invalid file extension: .{expected_extension} is expected"
            )
        return True

    def error_response(
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
        self._app.logger.error(msg)
        return {
            "success": False,
            "message": msg,
        }, code

    def _response(
        self,
        md_input_path: Path,
        docx_output_path: Path,
        output_formats: list[str],
        config_loader: ConfigLoader,
    ) -> Response:
        """Convert markdown resume to DOCX and optionally PDF

        Args:
            md_input_path (Path): Path to the input markdown file
            docx_output_path (Path): Path to the output DOCX file
            output_formats (list[str]): List of output formats
            config_loader (ConfigLoader): Configuration loader instance
            uuid_name (str): UUID of the input file (if applicable)

        Returns:
            Response: Flask response with the generated file
        """
        self._app.logger.info(f"Markdown input file: {md_input_path}")
        self._app.logger.info(f"Docx output file: {docx_output_path}")

        try:
            self._api_config.mimetypes.get("docx")
            self._api_config.mimetypes.get("pdf")

            # Convert markdown to DOCX
            docx_path = create_ats_resume(
                md_input_path,
                docx_output_path,
                config_loader=config_loader,
            )

            # Track created files and file to return
            output_file = None
            mime_types = None

            # Process DOCX if requested
            if DOCX_EXTENSION in output_formats:
                self._app.logger.info(f"Output extension: {DOCX_EXTENSION}")
                if os.path.exists(docx_path):
                    output_file = docx_path
                    mime_types = self._api_config.mimetypes.get("docx")

            # Process PDF if requested (convert from the generated DOCX)
            elif PDF_EXTENSION in output_formats:
                self._app.logger.info(f"Output extension: {PDF_EXTENSION}")
                pdf_path = convert_to_pdf(docx_path)
                if pdf_path and os.path.exists(pdf_path):
                    self._app.logger.info(f"PDF conversion successful: {pdf_path}")
                    output_file = pdf_path
                    mime_types = self._api_config.mimetypes.get("pdf")

            else:
                raise ValueError("Invalid output format specified")

            # If we don't have a file to return, that's an error
            self._app.logger.info(f"Output file: {output_file}")
            if not output_file or not os.path.exists(output_file):
                raise Exception(f"Failed to generate output file: {output_file}")

            # Return the appropriate file directly from the temp directory
            # Add explicit filename in Content-Disposition header for curl -O
            download_name = os.path.basename(output_file)
            self._app.logger.info(f"Successfully created: {output_file}")

            # Existing behavior - direct file download
            response = send_from_directory(
                directory=docx_output_path.parent,
                path=output_file.name,
                as_attachment=True,
                download_name=download_name,
                mimetype=mime_types[0],
            )

            # Force proper filename in Content-Disposition header
            response.headers["Content-Disposition"] = (
                f'attachment; filename="{download_name}"'
            )

            # Add additional headers to help browsers handle the download properly
            # response.headers["X-Content-Type-Options"] = "nosniff"

            return response

        except ValueError as e:
            return self.error_response(400, f"Value error: {str(e)}")
        except FileNotFoundError as e:
            self.error_response(404, e, "File not found")
        except Exception as e:
            return self.error_response(400, f"Error: {str(e)}")

    def post(
        self,
        output_format: str = DEFAULT_OUTPUT_FORMAT,
        request_body: str = None,
    ) -> Response | tuple[dict[str, Any], int]:
        """Convert markdown resume to DOCX and optionally PDF

        Args:
            output_format (str): Output format to generate (docx or pdf)
            request_body (str): Raw markdown content from request body

        Returns:
            Response: Flask response with the generated file
        """
        # Always load the config
        config_loader = ConfigLoader()

        # Get the uploaded file and parameters
        args = self._arg_parser.parse_args()
        input_file = args["input_file"]
        output_formats = (
            [output_format] if isinstance(output_format, str) else output_format
        )

        # Determine input source based on config and available inputs
        prefer_file = self._api_config.input.get("prefer_file", True)
        use_file_input = input_file is not None and (
            prefer_file or request_body is None
        )

        self._app.logger.info(f"Using file input: {use_file_input}")
        self._app.logger.info(f"Using request body: {request_body is not None}")
        self._app.logger.info(f"Request body: {request_body}")
        self._app.logger.info(f"Input file: {input_file}")

        if not use_file_input and not request_body:
            return self.error_response(
                400,
                ValueError("No input provided"),
                "Either input_file or request body must be provided",
            )

        # Get filename and output name
        if use_file_input:
            input_filename = Path(input_file.filename)
        else:
            random_id = uuid.uuid4().hex
            input_filename = Path(random_id).with_suffix(".md")

        base_output_filename = input_filename.stem
        output_name = f"{base_output_filename}.{DOCX_EXTENSION}"

        # Parse config_options if provided
        config_data = {}
        if args["config_options"]:
            try:
                config_data = json.loads(args["config_options"])
            except json.JSONDecodeError as e:
                return self.error_response(
                    400, e, "Invalid JSON in config_options parameter"
                )

        self._resolve_config_helper(config_loader, config_data)
        self._app.logger.debug(f"Configuration loaded: {config_loader.config}")
        temp_dir_enabled = self._api_config.output.get("use_temp_directory", True)
        self._app.logger.info(f"Temporary directory enabled: {temp_dir_enabled}")

        if temp_dir_enabled:
            with tempfile.TemporaryDirectory() as temp_dir:

                # Save the uploaded file
                temp_input_path = Path(temp_dir) / input_filename

                if use_file_input:
                    # Save the uploaded file
                    input_file.save(temp_input_path)
                else:
                    # Write the input text to a file, preserving non-UTF-8 characters
                    try:
                        # First try UTF-8
                        with open(temp_input_path, "w", encoding="utf-8") as f:
                            f.write(request_body)
                    except UnicodeEncodeError:
                        # If that fails, write binary
                        self._app.logger.info(
                            "UTF-8 encoding failed, writing as binary"
                        )
                        with open(temp_input_path, "wb") as f:
                            f.write(request_body.encode("utf-8", errors="replace"))

                # Prepare output paths directly in the temporary directory
                temp_output_path = Path(temp_dir) / output_name

                return self._response(
                    temp_input_path,
                    temp_output_path,
                    output_formats,
                    config_loader,
                )
        else:
            output_path = DEFAULT_OUTPUT_DIR / output_name
            return self._response(
                input_filename,
                output_path,
                output_formats,
                config_loader,
            )

    def _resolve_config_helper(
        self,
        config_loader: ConfigLoader,
        config_options: dict[str, Any] = None,
    ) -> None:
        """Merge the provided config options with the existing config

        Args:
            config_loader (ConfigLoader): Existing config loader
            config_options (dict): Configuration options to merge
        """
        if config_options:
            self._app.logger.info(f"Merging custom configuration: {config_options}")

            # Update top-level config sections
            for section_key, section_values in config_options.items():
                if section_key in config_loader.config:
                    # If section exists in default config, update it
                    if isinstance(section_values, dict) and isinstance(
                        config_loader.config[section_key], dict
                    ):
                        self._app.logger.debug(
                            f"Merging section '{section_key}' with values: {section_values}"
                        )
                        config_loader.config[section_key].update(section_values)
                    else:
                        # Replace the entire section if it's not a mergeable dictionary
                        self._app.logger.debug(
                            f"Replacing section '{section_key}' with values: {section_values}"
                        )
                        config_loader.config[section_key] = section_values
                else:
                    # Add new section if it doesn't exist
                    self._app.logger.debug(
                        f"Adding new section '{section_key}' with values: {section_values}"
                    )
                    config_loader.config[section_key] = section_values


app = App(SCRIPT_DIR / API_CONFIG_FILE)


@app.ns.route("/docx", methods=["POST"])
class ConvertDocxResource(Resource):
    @app.ns.doc(
        "convert_markdown",
        consumes=["text/plain", "multipart/form-data"],
    )
    @app.ns.expect(app.arg_parser)
    @app.ns.response(
        200,
        "Success - Returns DOCX file download",
    )
    @app.ns.response(
        400,
        "Bad Request",
        app.response_model,
        produces=app.api_config.mimetypes.get("error"),
    )
    @app.ns.response(
        404,
        "File Not Found",
        app.response_model,
        produces=app.api_config.mimetypes.get("error"),
    )
    @app.ns.response(
        500,
        "Server Error",
        app.response_model,
        produces=app.api_config.mimetypes.get("error"),
    )
    @app.ns.param(
        "payload",
        "Raw markdown content",
        _in="body",
        required=False,
        schema=App.TEXT_SCHEMA,
    )
    def post(self) -> Response:
        """Convert markdown resume to DOCX

        You can provide the markdown content either:
        - As a file upload (input_file)
        - Directly in the request body (Content-Type: text/plain)

        Returns:
            Response: Flask response with the generated DOCX file
        """
        content = request.get_data(as_text=True)
        return app.post(output_format=DOCX_EXTENSION, request_body=content)


@app.ns.route("/pdf", methods=["POST"])
class ConvertPdfResource(Resource):
    @app.ns.doc(
        "convert_markdown",
        consumes=["text/plain", "multipart/form-data"],
    )
    @app.ns.expect(app.arg_parser)
    @app.ns.response(
        200,
        "Success - Returns PDF file download",
    )
    @app.ns.response(
        400,
        "Bad Request",
        app.response_model,
        produces=app.api_config.mimetypes.get("error"),
    )
    @app.ns.response(
        404,
        "File Not Found",
        app.response_model,
        produces=app.api_config.mimetypes.get("error"),
    )
    @app.ns.response(
        500,
        "Server Error",
        app.response_model,
        produces=app.api_config.mimetypes.get("error"),
    )
    @app.ns.param(
        "body",
        "Raw markdown content",
        _in="body",
        required=False,
        schema=App.TEXT_SCHEMA,
    )
    def post(self) -> None:
        """Convert markdown resume to PDF

        You can provide the markdown content either:
        - As a file upload (input_file)
        - Directly in the request body (Content-Type: text/plain)

        Returns:
            Response: Flask response with the generated PDF file
        """
        content = request.get_data(as_text=True)
        return app.post(output_format=PDF_EXTENSION, request_body=content)


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
-F "config_options={\"document_styles\": {\"Subtitle\": {\"font_name\": \"Helvetica Neue\"}}}"
"""

    app.run(program_description, epilog_text)

# Export the Flask application object, not the App class instance
# This is what serverless-wsgi needs - the actual Flask application
application = app.app  # Get the Flask app from App class instance
