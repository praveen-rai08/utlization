"""
Setup configuration for Utilization Report Generator package
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="utilization-report-generator",
    version="1.0.0",
    author="QEA",
    description="Generate comprehensive employee leave and utilization reports",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/debjaniit/Utilization-Report.git",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Development Status :: 4 - Beta",
        "Intended Audience :: Business/Enterprise",
    ],
    python_requires=">=3.7",
    install_requires=[
        "click>=8.1.0",
        "openpyxl>=3.1.0",
        "Flask>=3.0.0",
        "Werkzeug>=3.0.0",
    ],
    entry_points={
        "console_scripts": [
            "utilization-report=utilization_report_generator.cli:cli",
        ],
    },
    include_package_data=True,
    package_data={
        'utilization_report_generator': [
            'templates/*',
            'static/*',
        ],
    },
)
