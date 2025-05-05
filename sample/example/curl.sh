API_KEY="${API_KEY:-$DEFAULT_API_KEY}"

curl -X POST "https://7lm0a3cnti.execute-api.us-east-1.amazonaws.com/dev/convert/docx" \
  -H "x-api-key: ${API_KEY}" \
  -H "Accept: application/vnd.openxmlformats-officedocument.wordprocessingml.document" \
  -d "$(cat ./sample/example/example.md)" -o resume.docx
