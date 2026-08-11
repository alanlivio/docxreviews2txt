.PHONY: deps test build publish-pypi clean

SHELL := /bin/bash

deps:
	pip install -e .
	pip install pytest build twine setuptools

test:
	pytest

build:
	python -m build . --wheel

publish-pypi-check:
	pip install build setuptools twine
	[[ -d dist ]] && rm -r dist || true
	[[ -d build ]] && rm -r build || true
	rm -rf ./*.egg-info
	python -m build . --wheel
	twine check dist/*

publish-pypi: publish-pypi-check
	twine upload dist/*

clean:
	rm -rf dist build ./*.egg-info .pytest_cache
