MAKEFLAGS += --no-print-directory

UNAME := $(shell uname -s)

.PHONY: all
all: init ## Entrypoint

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
init: .init .venv_reminder .python_command ## Ensure pip and Initialize venv

.PHONY: install
install: .install-api .venv_reminder .python_command ## Install dependencies

.PHONY: install-serverless
install-serverless: .install-npm .venv_reminder ## Install serverless dependencies

.PHONY: install-api
install-api: .install-api .venv_reminder ## Install api dependencies

.PHONY: install-dev
install-dev: .install-dev .venv_reminder .python_command ## Install development dependencies

.PHONY: uninstall
uninstall: .uninstall .venv_reminder .python_command ## Uninstall dependencies

.PHONY: uninstall-dev
uninstall-dev: .uninstall-dev .venv_reminder .python_command ## Uninstall development dependencies

.PHONY: build
build: .build ## Build the example
	@printf "Created %s from %s\n" "sample/template/output/sample.docx" "sample/template/sample.md"

.PHONY: deploy
deploy: .deploy-dev ## Deploy the dev environment

.PHONY: deploy-v1
deploy-v1: .deploy-v1 ## Deploy the v1 environment

.PHONY: remove
remove: .remove-dev ## Remove the dev environment

.PHONY: remove-v1
remove-v1: .remove-v1 ## Remove the v1 environment

.PHONY: clean
clean: .uninstall .uninstall-npm ## Clean up
	( \
  . .venv/bin/activate; \
  rm -rf .serverless src/__pycache__; \
  deactivate; \
  rm -rf .venv; \
)

.PHONY: check
check: ## Run linters but don't reformat
	( \
  . .venv/bin/activate; \
  black --check --diff . --line-length 88; \
  isort --check-only --diff .; \
  autoflake --check --remove-all-unused-imports --remove-unused-variables .; \
)

.PHONY: lint
lint: ## Run linters and reformat
	( \
  . .venv/bin/activate; \
  black . --line-length 88; \
  isort .; \
  autoflake --remove-all-unused-imports --remove-unused-variables .; \
)

.PHONY: api
api: ## Run the app
	( \
  . .venv/bin/activate; \
  python src/api.py --debug; \
)

.venv_reminder:
	@printf "\n\t📝 \033[1m%s\033[0m: %s\n\t   %s\n\t   %s\n\t   %s.\n\n\t🏄 %s \033[1;92m\`%s\`\033[0m\n\t   %s.\n" "NOTE" "The dependencies are installed" "in a virtual environment which needs" "to be manually activated to run the" "Python command" "Please run" ". .venv/bin/activate" "to activate the virtual environment"

.python_command:
	@printf "\n\033[1m%s\033[0m ...\n  \033[1;92m\`%s\`\033[0m\n  \033[1;92m\`%s\`\033[0m\n  \033[1;92m\`%s\`\033[0m\n\n" "The Python 🐍 command" "python3 resume_md_to_docx.py" "python3 resume_md_to_docx.py -i <input file>" "python3 resume_md_to_docx.py -i <input file> -o <output file>"

.init:
	@deactivate 2>/dev/null || true
	@test -d .venv || python3 -m venv .venv
	( \
  . .venv/bin/activate; \
  python3 -m ensurepip; \
)
	@printf "\nIf this is your \033[1m%s\033[0m running this (in this directory),\nplease \033[4m%s\033[0m\033[1m\033[0m run \033[1;92m\`%s\`\033[0m to install dependencies 🚀\n" "first time" "next" "make install"

.uninstall:
	( \
  . .venv/bin/activate; \
  pip uninstall -y -r src/requirements/requirements.txt; \
)

.uninstall-dev:
	( \
  . .venv/bin/activate; \
  pip uninstall -y -r src/requirements/requirements-dev.txt; \
  pre-commit uninstall; \
)

.install:
	( \
  . .venv/bin/activate; \
  pip install --no-cache-dir -r src/requirements/requirements.txt; \
)

.install-npm:
	( \
  . .venv/bin/activate; \
  sudo npm install -g serverless; \
  npm install serverless-wsgi serverless-python-requirements serverless-apigw-binary; \
)

.uninstall-npm:
	( \
  . .venv/bin/activate; \
  sudo npm uninstall -g serverless; \
  npm uninstall serverless-wsgi serverless-python-requirements serverless-apigw-binary; \
)

.install-api:
	( \
  . .venv/bin/activate; \
  pip install --no-cache-dir -r src/requirements/requirements-api.txt; \
)

.install-dev:
	( \
  . .venv/bin/activate; \
  pip install --no-cache-dir -r src/requirements/requirements-dev.txt; \
  pre-commit install; \
)

.build:
	( \
  . .venv/bin/activate; \
  python src/resume_md_to_docx.py -i sample/template/sample.md -o sample/template/output/sample.docx --pdf; \
  python src/resume_md_to_docx.py -i sample/template/sample.md -o sample/template/output/sample.paragraph-headings.docx -p h3 h4 h5 h6 --pdf; \
  python src/resume_md_to_docx.py -i sample/example/example.md -o sample/example/output/example.docx --pdf; \
  python src/resume_md_to_docx.py -i sample/example/example.md -o sample/example/output/example.paragraph-headings.docx -p h3 h4 h5 h6 --pdf; \
)

.deploy-dev:
	( \
  . .venv/bin/activate; \
  AWS_PROFILE=AdministratorAccess-254716123389 sls deploy; \
)

.deploy-v1:
	( \
  . .venv/bin/activate; \
  AWS_PROFILE=AdministratorAccess-254716123389 sls deploy --stage v1; \
)

.remove-dev:
	( \
  . .venv/bin/activate; \
  AWS_PROFILE=AdministratorAccess-254716123389 sls remove; \
)

.remove-v1:
	( \
  . .venv/bin/activate; \
  AWS_PROFILE=AdministratorAccess-254716123389 sls remove --stage v1; \
)
