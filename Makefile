MAKEFLAGS += -s --no-print-directory
.DEFAULT_GOAL := help

.PHONY: help deps test build clean wheel publish-pypi

help:
	@printf "%s\n" \
		"Usage: make [target]" \
		"" \
		"Targets:" \
		"  deps          Install dependencies" \
		"  test          Run tests" \
		"  build         Build wheel" \
		"  wheel         Build and check wheel" \
		"  publish-pypi  Publish wheel to PyPI" \
		"  clean         Clean build artifacts"

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

publish-pypi: wheel
	twine upload dist/*


