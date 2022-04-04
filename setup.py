import setuptools

with open('requirements.txt') as f:
    required = f.read().splitlines()

setuptools.setup(
    name="docx_reviews_to_txt",
    version="0.1",
    author="Alan Guedes",
    license='MIT',
    url="http://github.com/alanlivio/docx_reviews_to_txt",
    python_requires='>= 3',
    install_requires=required,
    classifiers=[
        "Development Status :: 1 - Planning",
        "Environment :: Console",
        "Operating System :: Microsoft :: Windows"
        "Intended Audience :: Science/Research",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3", ],
    author_email="alanlivio@gmail.com",
    description="Extract reviews changes and commentaries from a docx file as plan text.",
    scripts=['src/docx_reviews_to_txt.py'],
)
