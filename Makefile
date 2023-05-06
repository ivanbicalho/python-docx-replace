.PHONY: install
install:
	pip install -r requirements/dev.txt -r requirements/prod.txt

.PHONY: venv
venv:
	python3 -m venv .venv

.PHONY: black
black:
	black --line-length 120 src

.PHONY: ruff
ruff:
	ruff check src --fix

.PHONY: mypy
mypy:
	mypy src

.PHONY: fix
fix: black ruff mypy
