MAKEFLAGS += -s --no-print-directory
.DEFAULT_GOAL := help

GLOBAL_PYTHON ?= $(shell if [ -x /usr/bin/python3 ]; then echo /usr/bin/python3; else echo python3; fi)

.PHONY: help deps test build clean wheel install-global publish-pypi

help:
	@printf "%s\n" \
		"Usage: make [target]" \
		"" \
		"Targets:" \
		"  deps            Install dependencies" \
		"  test            Run tests" \
		"  build           Build wheel" \
		"  clean           Clean build artifacts" \
		"  wheel           Build and check wheel" \
		"  install-global  Install built wheel globally" \
		"  publish-pypi    Publish wheel to PyPI"

deps:
	pip install --upgrade pip
	pip install -r requirements.txt -r requirements-dev.txt

test:
	pytest

build:
	python -m build . --wheel

clean:
	rm -rf dist build ./*.egg-info .pytest_cache

wheel:
	pip install -r requirements-dev.txt
	rm -rf dist build ./*.egg-info
	python -m build . --wheel
	twine check dist/*

install-global: wheel
	$(GLOBAL_PYTHON) -m pip install --force-reinstall --break-system-packages dist/*.whl

publish-pypi: wheel
	twine upload dist/*


