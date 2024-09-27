from setuptools import setup, find_packages
from pathlib import Path

# Read the requirements from the requirements.txt file
with open('requirements.txt', encoding='utf-8') as f:
    requirements = f.read().splitlines()

# Read the README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding="utf-8")

setup(
    name='excel_data_pairer',  
    version='0.0.1',  # Initial version
    author='Lael Al-Halawani',
    author_email='laelhalawani@gmail.com',
    description='Get data pairs from excel sheets, easily and automatically build jsons. Perfect for extracting translations from excel files.',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='https://github.com/laelhalawani/excel_data_pairer',
    packages=find_packages(),
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.10',  # Changed to >=3.10 instead of 3.10.14
    install_requires=requirements,
)