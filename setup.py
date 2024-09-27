from setuptools import setup, find_packages

setup(
    name='excel_navigator',  # Replace with your package name
    version='0.0.1',  # Initial version
    author='Lael Al-Halawani',
    author_email='laelhalawani@gmail.com',
    description='Get',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/yourusername/your-repo',  # Replace with your repo URL
    packages=find_packages(),
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
    install_requires=[
        # List your package dependencies here
        # 'some_package>=1.0.0',
    ],
)