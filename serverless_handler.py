from serverless_wsgi import handle_request

BASE64_DECODE = False


def handler(event, context):
    """WSGI handler for API Gateway binary responses"""
    try:
        # Get API response
        from src.api import application

        response = handle_request(application, event, context)
        headers = response.get("headers", {})

        # Check if response is base64-encoded but not JSON
        if (
            BASE64_DECODE
            and response.get("isBase64Encoded")
            and headers.get("Content-Type") != "application/json"
        ):
            import base64

            print(f"Found base64-encoded binary response - decoding it")

            # Get the encoded content
            encoded_body = response.get("body", "")

            # Convert the binary data to a string with latin-1 encoding
            # This preserves all byte values when marshaling through Lambda
            decoded_body = base64.b64encode(encoded_body).decode("latin1")

            # Return the decoded content
            return {
                "statusCode": response.get("statusCode", 200),
                "headers": headers,
                "body": decoded_body,
                "isBase64Encoded": True,  # Important: Set to false since we've already decoded it
            }

        # For all other responses (JSON, non-binary, etc.), pass through unchanged
        return response
    except Exception as e:
        import json
        import traceback

        print(f"ERROR: {str(e)}")
        print(traceback.format_exc())
        return {
            "statusCode": 500,
            "headers": {"Content-Type": "application/json"},
            "body": json.dumps({"error": str(e)}),
        }
