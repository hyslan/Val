'''
    Author: Hyslan Silva Cruz
    Built for work by Hyslan on SAP GUi v7.40
    Client: SABESP
    
'''
from setuptools import setup, find_packages

setup(
    name='VAL',
    version='0.1.0',
    packages=find_packages(),
    install_requires=[
        "numpy", "pandas", "openpyxl", "pywin32",
        "rich", "sqlalchemy", "selenium", "tqdm",
        "imageio", "pygame", "dotenv"
    ],
    entry_points={
        'console_scripts': [
            'val=src.main:main',
        ],
    },
    author='Hyslan Silva Cruz',
    author_email='hyslansilva@gmail.com',
    description='Valoração automática de ordens',
    long_description=open('README.md', encoding='latin1').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/hyslan/Val',
    classifiers=[
        'Programming Language :: Python :: 3.11.4',
        'Operating System :: OS Independent',
    ],
)
