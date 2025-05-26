from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="fp4nt",
    version="1.1.2",
    author="charlie3go",
    author_email="aslongrushan@gmail.com",
    description="Fp4NT is a file parsing tool specially designed for processing and converting nested tabular data. It supports multiple input formats, such as HTML and Excel, and converts them to Markdown-formatted tables.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/charlie3go/fp4nt",
    packages=find_packages(),
    install_requires=[
        'beautifulsoup4>=4.9.0',
        'html2text>=2020.1.16',
        "pandas>=1.0.0",
        "openpyxl>=3.0.0"
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.7',
)
