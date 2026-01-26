"""
Setup configuration for docx-template-export package.
"""
from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text(encoding="utf-8") if readme_file.exists() else ""

setup(
    name="docx-template-export",
    version="0.3.1",
    description="A deterministic library for exporting markdown content to Word (.docx) templates",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Ahsan Saeed",
    url="https://github.com/asaeed9/deterministic-docx-export",
    packages=find_packages(exclude=["tests", "tests.*"]),
    include_package_data=True,
    python_requires=">=3.8",
    install_requires=[
        "pydantic>=2.0.0",
        "python-docx>=1.1.0",
        "markdown-it-py>=3.0.0",
    ],
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: Apache Software License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Text Processing :: Markup",
    ],
    keywords="docx word markdown export template deterministic",
)
