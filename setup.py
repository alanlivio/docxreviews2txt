import setuptools
from docxreviews2txt.docxreviews2txt import __version__

with open("requirements.txt") as f:
    required = f.read().splitlines()

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="docxreviews2txt",
    packages=["docxreviews2txt"],
    version=__version__,
    author="Alan Guedes",
    license="MIT",
  url="http://github.com/alanlivio/docxreviews2txt",
    python_requires=">= 3.6",
    install_requires=required,
    long_description=long_description,
    long_description_content_type="text/markdown",
    classifiers=[
        "Development Status :: 1 - Planning",
        "Environment :: Console",
        "Operating System :: Microsoft :: Windows",
        "Intended Audience :: Science/Research",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
    ],
    author_email="alanlivio@gmail.com",
    description="Command line tool to extract review changes and comments from a docx file as plain text.",
    entry_points={
        "console_scripts": [
            "docxreviews2txt = docxreviews2txt.__main__:main",
        ]
    },
)
