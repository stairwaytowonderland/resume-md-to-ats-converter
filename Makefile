MAKEFLAGS += --no-print-directory

UNAME := $(shell uname -s)
SCRIPT_DIR := $(shell dirname $(realpath $(lastword $(MAKEFILE_LIST))))

.PHONY: all
all: .init .venv_reminder .python_command ## Entrypoint

.PHONY: help
help: ## Show this help.
	@echo "Please use \`make <target>' where <target> is one of"
	@grep '^[a-zA-Z]' $(MAKEFILE_LIST) | \
    sort | \
    awk -F ':.*?## ' 'NF==2 {printf "\033[36m  %-26s\033[0m %s\n", $$1, $$2}'

.list-targets:
	@LC_ALL=C $(MAKE) -pRrq -f $(firstword $(MAKEFILE_LIST)) : 2>/dev/null | awk -v RS= -F: '/(^|\n)# Files(\n|$$)/,/(^|\n)# Finished Make data base/ {if ($$1 !~ "^[#.]") {print $$1}}' | sort

.PHONY: list
list: ## List public targets
	@LC_ALL=C $(MAKE) .list-targets | grep -E -v -e '^[^[:alnum:]]' -e '^$@$$' | xargs -n3 printf "%-26s%-26s%-26s%s\n"

.PHONY: init
init: .init .venv_reminder .python_command ## Install dependencies

.PHONY: install
install: .install .venv_reminder .python_command ## Install dependencies

.PHONY: clean
clean: .uninstall ## Clean up
	@( \
  . .venv/bin/activate; \
  deactivate; \
  rm -rf .venv; \
)

.PHONY: test
test: ## Run linters but don't reformat
	@( \
  . .venv/bin/activate; \
  black --check --diff resume_md_to_docx.py; \
  isort --check-only --diff resume_md_to_docx.py; \
  autoflake --check --remove-all-unused-imports --remove-unused-variables resume_md_to_docx.py; \
)

.PHONY: lint
lint: ## Run linters and reformat
	@( \
  . .venv/bin/activate; \
  black resume_md_to_docx.py; \
  isort resume_md_to_docx.py; \
  autoflake --remove-all-unused-imports --remove-unused-variables resume_md_to_docx.py; \
)

.venv_reminder:
	@printf "\n\t📝 \033[1m%s\033[0m: %s\n\t   %s\n\t   %s\n\t   %s.\n\n\t🏄 %s \033[1;92m\`%s\`\033[0m\n\t   %s.\n" "NOTE" "The dependencies are installed" "in a virtual environment which needs" "to be manually activated to run the" "Python command" "Please run" ". .venv/bin/activate" "to activate the virtual environment"

.python_command:
	@printf "\n\033[1m%s\033[0m (%s) 🐍 \n  \033[1;92m\`%s\`\033[0m # for usage help\n  \033[1;92m\`%s\`\033[0m\n  \033[1;92m\`%s\`\033[0m\n\n" "The Python command(s)" "you must manually activate the virtual environment" "python3 resume_md_to_docx.py" "python3 resume_md_to_docx.py -i <input file>" "python3 resume_md_to_docx.py -i <input file> -o <output file>"

.init:
	@deactivate 2>/dev/null || true
	@test -d .venv || python3 -m venv .venv
	@( \
  . .venv/bin/activate; \
  python3 -m ensurepip; \
)
	@printf "\nIf this is your \033[1m%s\033[0m running this (in this directory),\nplease \033[4m%s\033[0m\033[1m\033[0m run \033[1;92m\`%s\`\033[0m to install dependencies 🚀\n" "first time" "next" "make install"

.uninstall:
	@( \
  . .venv/bin/activate; \
  pip uninstall -y python-docx markdown beautifulsoup4 black isort autoflake; \
)

.install:
	@( \
  . .venv/bin/activate; \
  pip install --no-cache-dir python-docx markdown beautifulsoup4 black isort autoflake; \
)
