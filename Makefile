.PHONY: install
install:
	@pip install -r requirements/dev.txt

.PHONY: venv
venv:
	@python3 -m venv .venv

.PHONY: fix
fix:
	@echo "==> fixing black"
	@black --line-length 120 src/

	@echo "==> fixing isort"
	@isort src/

.PHONY: check
check:
	@echo "==> checking isort"
	@isort --check-only --diff src

	@echo "==> checking black"
	@black --check --diff --line-length 120 src

	@echo "==> checking flake8"
	@flake8 src

	@echo "==> checking mypy"
	@mypy src