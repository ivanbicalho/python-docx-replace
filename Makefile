.PHONY: install
install:
	@pip install -r requirements/dev.txt

.PHONY: venv
venv:
	@python3 -m venv .venv
	@source .venv/bin/activate

.PHONY: fix
fix:
	@isort src
	@black src
	@flake8 src