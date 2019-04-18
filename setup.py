from pathlib import Path
from setuptools import setup, find_packages


with open(Path('./README.rst').resolve()) as f:
    readme = f.read()

with open(Path('LICENSE').resolve()) as f:
    license_text = f.read()

setup(
    name='diff_excel',
    version='0.1.0',
    description='2つのエクセルデータを比較する',
    entry_points={
        "console_scripts": [
            "diff_excel=diff_excel.__init__:main"
        ]
    },
    long_description=readme,
    author='tthmmts',
    author_email='tthmmts@gmail.com',
    url='https://github.com/tthmmts',
    license=license_text,
    packages=find_packages(exclude=('tests', 'docs')),
    test_suite='tests'
)

