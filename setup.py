import setuptools

with open('requirements.txt') as f:
    required = f.read().splitlines()

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="docxreviews2txt",
    version="0.1",
    author="Alan Guedes",
    license='MIT',
    url="http://github.com/alanlivio/docxreviews2txt",
    python_requires='>= 3.6',
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
        "Programming Language :: Python :: 3", ],
    author_email="alanlivio@gmail.com",
    description="Extract reviews changes and commentaries from a docx file as plan text.",
    scripts=['scripts/docxreviews2txt'],
)
