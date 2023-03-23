from setuptools import setup, find_packages


setup(
    name="allure-docx",
    description="docx report generator based on allure-generated json files",
    author="Victor Maryama (Typhoon HIL, Inc)",
    version="0.4.0",
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
)
