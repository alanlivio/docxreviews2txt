# Build and Run Guide

Instructions for setting up, building, running from source, and running tests for docxreviews2txt.

## Environment Setup

Create and activate a virtual environment:

```bash
# Windows (PowerShell)
python -m venv .venv
. .venv\Scripts\Activate.ps1

# Linux / macOS (Bash)
python3 -m venv .venv
source .venv/bin/activate
```

## Install Dependencies

```bash
make deps
```

## Run from Source

Run directly using Python module execution without installing globally:

```bash
python -m docxreviews2txt tests/input_1.docx
```

Or run the CLI command directly when installed in editable mode:

```bash
docxreviews2txt tests/input_1.docx
```

Specify output formats (`diff` or `tags`):

```bash
# Output with PREVIOUS -> AFTER diff style (default)
docxreviews2txt --format diff tests/input_1.docx

# Output with <ins> and <del> tags
docxreviews2txt --format tags tests/input_1.docx
```

Display help and usage details:

```bash
python -m docxreviews2txt --help
```

## Run Tests

Run the test suite using `make`:

```bash
make test
```

## Build Distribution Packages

Build wheel and distribution packages using `make`:

```bash
make wheel
```

Built package artifacts will be created in the `dist` directory.

To clean build artifacts:

```bash
make clean
```
