# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = 'Automation Correlation and Cpk Report'
copyright = '2024, Mohamad Haikal bin Mohamad Nazari'
author = 'Mohamad Haikal bin Mohamad Nazari'
release = '1.0'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = ['nbsphinx', 'myst_parser', "sphinxcontrib.drawio"]

templates_path = ['_templates']
exclude_patterns = []

source_suffix = ['.rst','.txt','.md']
master_doc = 'index'

# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'sphinx_rtd_theme'
html_static_path = ['_static']
