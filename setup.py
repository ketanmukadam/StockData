# -*- coding: utf-8 -*-

from setuptools import setup, find_packages


with open('README.md') as f:
    readme = f.read()

with open('LICENSE.md') as f:
    license = f.read()

setup(
    name='stockdata',
    version='0.0.1',
    description='Fetch Indian Stock Financial Data',
    long_description=readme,
    author='Ketan Mukadam',
    author_email='ketanm@gmail.com',
    url='https://github.com/ketanmukadam/StockData',
    license=license,
    packages=find_packages(exclude=('tests', 'docs'))
)

