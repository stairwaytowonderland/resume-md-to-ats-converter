#!/usr/bin/env bash

set -eu

__python_module="src.resume_md_to_docx"

__script_dir() {
  local src=""
  local dir=""
  if [ -z "${BASH_SOURCE[0]}" ] || [ "${BASH_SOURCE[0]}" != "$0" ]; then
    src="${0}/bin"
  else
    src="${BASH_SOURCE[0]}"
  fi
  dir=$(CDPATH='' cd -- "$( dirname -- "$src" )" &> /dev/null && pwd -P)
  echo "$dir"
}

__python_usage() {
  local code="${1:-0}"
  python -m "$__python_module" --help
  [ "$code" -eq 0 ] || exit "$code"
}

__usage() {
  local code="${1:-0}"
  cat <<EOF

$(__python_usage)

Wrapper Script ...
Usage: $(basename "$0") <markdown_file> <output_directory> [additional_args]

Arguments:
  <markdown_file>     Path to the Markdown file to convert.
  <output_directory>  Directory where the output DOCX file will be saved.
  [additional_args]   Additional arguments to pass to the conversion script.

Example:
  $0 sample/example/example.md sample/example/output
EOF
  [ "$code" -eq 0 ] || exit "$code"
}

__usage_check() {
  if [ "$#" -lt 2 ]; then
    printf "Wrapper Script Error: Insufficient arguments provided ...\n \
 ... Received '%d', Expected at least '%d'\n" "$#" "2" 1>&2
    [ "$#" -eq 0 ] || __usage
    return 1
  fi
}

__file_exists() {
  local file="$1"
  if [ ! -r "$file" ]; then
    echo "File $file does not exist or is not readable."
    return 1
  fi
}

__command() {
  local module="$1"; shift
  local dir="$1"; shift
  pushd "$dir" > /dev/null || return 1
  local md_file="$1"; shift
  local output_dir="$1"; shift
  local file_name
  file_name=$(basename "$md_file")
  local base_file_name="${file_name%.*}"
  local output_file="${output_dir}/${base_file_name}.docx"
  __usage_check "$md_file" "$output_dir" || return 1
  __file_exists "$md_file" || return 1
  if ! __file_exists "$output_dir"; then
    printf "Creating output directory: '%s'\n" "$output_dir"
    mkdir -p "$output_dir"
  fi
  printf "Generating DOCX from Markdown file: '%s'\n" "$(pwd)/$md_file"
  printf "Output will be saved to: '%s'\n" "$(pwd)/$output_file"
  set -x
  python -m "$module" --pdf -i "$md_file" -o "$output_file" "$@"
}

main() {
  local dir
  dir=$(__script_dir)
  if [ ! -r "$dir" ]; then
    printf "Directory '%s' does not exist or is not readable.\n" "$dir"
    return 1
  fi
  __command "$__python_module" "${dir}/.." "$@" || return 1
  set +x
  popd > /dev/null || return 1
}

__usage_check "$@" || exit 1

main "$@"

# Final error check
if [ "$?" -ne 0 ]; then
  echo "An error occurred during the execution of the script."
  exit 1
fi

# Exit gracefully
set +e
set +u
