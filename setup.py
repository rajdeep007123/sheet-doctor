from setuptools import setup


setup(
    name="sheet-doctor",
    version="0.5.0",
    description="Free local spreadsheet diagnosis and cleanup tools for messy CSV and Excel files",
    packages=["sheet_doctor"],
    package_data={
        "sheet_doctor": [
            "bundled/skills/csv-doctor/scripts/*.py",
            "bundled/skills/csv-doctor/scripts/heal_modules/*.py",
            "bundled/skills/excel-doctor/scripts/*.py",
        ]
    },
    include_package_data=True,
    install_requires=[
        "pandas",
        "chardet",
        "openpyxl",
        "streamlit",
        "requests",
    ],
    extras_require={
        "excel-legacy": ["xlrd"],
        "ods": ["odfpy"],
        "all": ["xlrd", "odfpy"],
    },
    entry_points={
        "console_scripts": [
            "sheet-doctor=sheet_doctor.cli:main",
        ]
    },
)
