.PHONY: deps test build publish-pypi clean

SHELL := /bin/bash

deps:
	pip install -e .
	pip install pytest build twine setuptools

test:
	pytest

build:
	python -m build . --wheel

clean:
	rm -rf dist build ./*.egg-info .pytest_cache

wheel:
	$(VENV)/bin/pip install build setuptools twine
	rm -rf dist build ./*.egg-info
	$(VENV)/bin/python -m build . --wheel
	$(VENV)/bin/twine check dist/*

publish-pypi: wheel
	twine upload dist/*
