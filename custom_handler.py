import json
import base64
from serverless_wsgi import handle_request


def handler(event, context):
    """WSGI handler for API Gateway binary responses"""
    try:
        # Get API response
        from src.api import application
        response = handle_request(application, event, context)

        return response
    except Exception as e:
        import traceback
        print(f"ERROR: {str(e)}")
        print(traceback.format_exc())
        return {
            'statusCode': 500,
            'headers': {'Content-Type': 'application/json'},
            'body': json.dumps({"error": str(e)}),
        }
