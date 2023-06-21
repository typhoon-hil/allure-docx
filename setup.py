from setuptools import setup, find_packages
from pathlib import Path

this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text()

setup(
    name="allure-docx",
    description="Docx report generator based on allure-generated json files.",
    author="Victor Maryama (Typhoon HIL, Inc), Bjarne Scheimann (Fraunhofer Institute for Solar Energy Systems ISE)",
    version="0.4.0a2",
    license="MIT",
    install_requires=[
        'setuptools-git~=1.2',
        'matplotlib>=3.0, < 4.0',
        'docx2pdf~=0.1.8',
        'click',
        'python-docx',
    ],
    extras_require={
        'dev': ['pyinstaller'],
    },

    packages=find_packages('src'),
    package_dir={'': 'src'},
    # Should be present so MANIFEST.in is taken into account. However, only adds files that are inside package.
    include_package_data=True,

    entry_points={
        'console_scripts': ['allure-docx = allure_docx.commandline:main'],
    },
    long_description=long_description,
    long_description_content_type='text/markdown',

    project_urls={
        'Source':  'https://github.com/typhoon-hil/allure-docx',
        'Tracker': 'https://github.com/typhoon-hil/allure-docx/issues',
    },
)
