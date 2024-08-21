"""Author: Hyslan Silva Cruz
Built for work by Hyslan on SAP GUi v7.40
Client: SABESP

"""

with open("requirements.txt") as f:
    required = f.read().splitlines()

from setuptools import find_packages, setup

setup(
    name="VAL",
    version="0.2.0",
    packages=find_packages(),
    install_requires=required,
    entry_points={
        "console_scripts": [
            "val=src.main:main",
        ],
    },
    author="Hyslan Silva Cruz",
    author_email="hyslansilva@gmail.com",
    description="Valoração automática de ordens dentro do SAPGUI v7.40",
    long_description=open("README.md", encoding="latin1").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/hyslan/Val",
    classifiers=[
        "Programming Language :: Python :: 3.11.4",
        "Operating System :: OS Independent",
    ],
)
