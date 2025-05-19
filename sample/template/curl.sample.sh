API_KEY="${API_KEY:-$DEFAULT_API_KEY}"

curl -X POST "https://7lm0a3cnti.execute-api.us-east-1.amazonaws.com/dev/convert/docx" \
  -H "x-api-key: ${API_KEY}" \
  -H "Accept: application/octet-stream" \
  -d "$(cat ./sample/template/sample.md)" -o resume.docx
