# -*- coding: utf-8 -*-
"""
setup.py
~~~~~~~~~~~~~~
MathcadPy
Copyright 2024 Matt Woodhead
"""

from pathlib import Path
from setuptools import setup

# get key package details from MathcadPy/__version__.py
about = {}  # type: ignore
here = Path(__file__).parent
with open(here / 'MathcadPy' / '__version__.py') as f:
    exec(f.read(), about)

# load the README file and use it as the long_description for PyPI
with open('README.md', 'r') as f:
    readme = f.read()

# package configuration - for reference see:
# https://setuptools.readthedocs.io/en/latest/setuptools.html#id9
setup(
    name=about['__title__'],
    description=about['__description__'],
    long_description=readme,
    long_description_content_type='text/markdown',
    version=about['__version__'],
    author=about['__author__'],
    author_email=about['__author_email__'],
    url=about['__url__'],
    packages=['MathcadPy'],
    include_package_data=True,
    python_requires=">3.5",
    install_requires=['pywin32'],
    license=about['__license__'],
    zip_safe=False,
    entry_points={
        'console_scripts': ['py-package-template=py_pkg.entry_points:main'],
    },
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'Programming Language :: Python :: 3.6',
        'Environment :: Win32 (MS Windows)',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
    ],
    keywords='Mathcad, automation, COM, windows'
)
