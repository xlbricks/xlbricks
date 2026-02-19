"""
Setup configuration for xlbricks package.
"""

from setuptools import setup, find_packages
import os

# Read the long description from README
here = os.path.abspath(os.path.dirname(__file__))
readme_path = os.path.join(here, 'README.md')

long_description = ''
if os.path.exists(readme_path):
    with open(readme_path, 'r', encoding='utf-8') as f:
        long_description = f.read()

setup(
    name='xlbricks',
    version='0.1.0',
    author='julij.jegorov',
    author_email='',
    description='Excel-integrated brick structures and User Defined Functions (UDFs)',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='',
    packages=find_packages(exclude=['tests', '*.tests', '*.tests.*', 'tests.*']),
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'Intended Audience :: Financial and Insurance Industry',
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Topic :: Office/Business :: Financial',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    python_requires='>=3.7',
    install_requires=[
        'numpy>=1.19.0',
        'pandas>=1.1.0',
        'xlwings>=0.24.0',
        'PyQt5>=5.15.0',
        'QuantLib-Python>=1.26',
    ],
    extras_require={
        'dev': [
            'pytest>=6.0.0',
            'pytest-cov>=2.10.0',
            'build>=0.7.0',
            'twine>=3.4.0',
        ],
    },
    package_data={
        'xlbricks': ['xlbricks.json'],
    },
    include_package_data=True,
    keywords='excel xlwings udf quantlib finance pyqt5',
    project_urls={
        'Bug Reports': '',
        'Source': '',
    },
)
