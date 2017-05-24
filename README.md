# LinguisticTools
Collection of Linguistic Tools

LinguisticTools:
- GlossaryCheck (gc58.py)
- MySQLCheck002 (MySQLCheck002.py)
- Search+Mark (SAM39.py)

GlossaryCheck is a linguistic tool to help find terminology errors in large string based localization projects using spreadsheet files. GlossaryCheck works with bulk inputs of glossaries (up to 1048576 entries) and string files (up to 1048576 strings) via .xlsx spreadsheets. GlossaryCheck is independent of other Computer-assisted translation software and can be used as an analyzing tool, analyzing the bulk output of several translators, collected in one single spreadsheet (e.g. selective output of large MySql databases). The output of GlossaryCheck is a .xlsx spreadsheet with detailed info about localization terminology errors (String ID, terminology term in source and localized language, original string, localized string). GlossaryCheck's input tool allows different sensitivity settings e.g. case insensitive/sensitive, word boundaries insensitive/sensitive for thousands of Glossary terms at the same time. GlossaryCheck contains also a tool to create or extend terminology lists (GCCreator). 

MySQLCheck is a tool to search for multiple terms (up to 1048576) in strings of MySQL databases. MySQLCheck works with bulk inputs of term collections (up to 1048576 entries)  via .xlsx spreadsheets. The output of MySQLCheck is a .xlsx spreadsheet with data about the found strings (ID, string, term, corresponding term and corresponding string).

Search+Mark is a tool to search for and mark entries in multiple .xlsx spreadsheet files. Different settings allow to adjust the sensitivity of the search engine as well as the location, cell color and content of the mark and comment input. Search+Mark allows to assign different selections of multiple .xlsx spreadsheets and to store this selections permanent. Search+Mark informs the user about all important data, without the need to open the .xlsx file. Search+Mark contains the editor tool Search+Edit. With Search+Edit it is possible to search in multiple .xlsx spreadsheet files and to edit any cell, without opening the spreadsheet in Office.

Search+Mark is written by A.D.Klumpp using Python and the Python library openpyxl including jdcal and et_xmlfile (see license texts below or in the folders of the libraries). Search+Mark is released under the terms of the GNU General Public License (See http://www.gnu.org/licenses/). Copyright (C) 2016 A.D.Klumpp. Search+Mark is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY. The full copyright notices and the full license texts shall be included in all copies or substantial portions of the Software.

MySQLCheck is based on GlossaryCheck and is written by A.D.Klumpp using Python and the Python library openpyxl including jdcal and et_xmlfile (see license texts below or in the folders of the libraries). MySQLCheck/GlossaryCheck is released under the terms of the GNU General Public License (See http://www.gnu.org/licenses/).

GlossaryCheck is written by A.D.Klumpp using Python and the Python library openpyxl including jdcal and et_xmlfile (see license texts below or in the folders of the libraries). GlossaryCheck is released under the terms of the GNU General Public License (See http://www.gnu.org/licenses/). Copyright (C) 2015 A.D.Klumpp. GlossaryCheck is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY. The full copyright notices and the full license texts shall be included in all copies or substantial portions of the Software. Alternative versions of GlossaryCheck are written by A.D. Klumpp using Python and the Python libraries openpyxl (including jdcal and et_xmlfile), XlsxWriter and kivy (thanks to python-for-android). 

Python is released under the Python Software Foundation License (see https://www.python.org/download/releases/2.7.6/license/). Openpyxl is released under MIT/Expat license (see https://openpyxl.readthedocs.org/en/latest/). Kivy is released under the MIT License (see https://github.com/kivy/kivy/blob/master/LICENSE). jdcal is released under BSD (see https://pypi.python.org/pypi/jdcal). et_xmlfile is released under MIT (Home-page: https://bitbucket.org/openpyxl/et_xmlfile).

Please read the full license texts Online or in the folder.

Quick start:
1) Install Python 2.7 (full installation, including "Add python.exe to search path")
2) Run GlossaryCheck (gc58.py) or MySQLCheck002 (MySQLCheck002.py) or Search+Mark (SAM39.py). [Set this files as executable, use always: 'run in terminal'].
3) Read the manual (START-menu).




