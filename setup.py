from setuptools import setup, find_packages

# Read the requirements from the requirements.txt file
with open('requirements.txt') as f:
    requirements = f.read().splitlines()

setup(
    name='excel_navigator',  # Replace with your package name
    version='0.0.1',  # Initial version
    author='Lael Al-Halawani',
    author_email='laelhalawani@gmail.com',
    description='Get data pairs from excel sheets, easily and automatically build jsons. Perfect for extracting translations from excel files.',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/laelhalawani/excel_data_pairer',  # Replace with your repo URL
    packages=find_packages(),
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.10.14',
    install_requires=requirements,
)