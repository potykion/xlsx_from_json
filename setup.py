import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

requirements = [
    "attrs>=18.1.0",
    "openpyxl>=2.5.4"
]

setuptools.setup(
    name="xlsx_from_json",
    version="0.3.0",
    author="potykion",
    author_email="potykion@gmail.com",
    description="Creates xlsx from json via openpyxl.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/potykion/xlsx_from_json",
    install_requires=requirements,
    packages=setuptools.find_packages(),
    classifiers=(
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ),
    include_package_data=True,
)
