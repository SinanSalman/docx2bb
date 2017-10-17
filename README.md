# README #

## OVERVIEW ##
**docx2bb** is a tool for creating BlackBoard (bb) test questions (text) import file from a Microsoft Word \*.docx document. Supported question types are: True/False, Multiple choice, Matching, Essay, and (simple) Fill in the blank. ExamFormat-Sample.docx shows a sample exam format for use with docx2bb. Unicode-to-ASCII replacement rules from 'docx2bb.json' data file can be optionally applied.

docx2bb includes three components:

*   docx2bb_lib - library including the conversion logic
*   docx2bb     - command line interface (cli)
*   docx2bb_web - website interface

An important design principle for docx2bb was that it's input (MS-word docx) file must not look different from a key solution exam. this way a specifically prepared exam key solution document can be processed by docx2bb and a import text file results, reducing the number of steps and files necessary to manage the exam automation process.

'''
Syntax:
   python docx2bb.py \[options\] \[docx_filename\]
options:
   --verbose | -v display verbose messages
   --help | -h display help message
'''

## Installation ##
**docx2bb** can be installed as a python3 package:

*   Clone the project's repository or download its zip file from [github.com](https://sinansalman.github.io/docx2bb/)
*   Unzip the file to a folder on your hard drive and rename the resulting folder to 'docx2bb'
*   Install the python package and start it

```
git clone https://github.com/SinanSalman/docx2bb.git
cd docx2bb
pip install .
./srart.sh
```

## SOURCE CODE ##
The source distribution contains Python, JavaScript, CSS, HTML code. The code also makes use of several libraries including Python-Flask, jQuery, and python-docx.

## CONTRIBUTE	##
Code submissions are greatly appreciated and highly encouraged. Please send fixes, enhancements, etc. to SinanSalman at GitHub or sinan\[dot\]salman\[at\]zu\[dot\]ac\[dot\]ae.

## LICENSE	##
docx2bb is released under the GPLv3 license, which is available at [GNU](https://www.gnu.org/licenses/gpl-3.0.en.html)

## Disclaimer ##
**docx2bb** is provided with no warranties, use it if you find it useful. docx2bb is designed to keep your \*.docx document unchanged, but the author assumes no liabilities from use of this tool, including if it eats your exam ;).

## COPYRIGHT ##
2016-2017 Sinan Salman, PhD

## Version and History ##
*   Sep 28th, 2017	0.20	Initial web interface release
*   Sep 20th, 2017 0.17  fixed bug with single answer FIB identified as ESS
*   Mar 07th, 2017	0.16	added tabs and ... to the list of replaced characters
*   Feb 13th, 2017	0.15	fixed Q beg/end bug
*   Nov 29th, 2016	0.14	fixed minor py2.7 file open compatibility issue (JSON). Cleanup in prep for pyinstall packaging
*   Nov 17th, 2016	0.13	add default values for unicode2ascii replacements even if docx2bb.jason file is absent
*   Nov 07th, 2016	0.12	add PageBreak elimination logic, fix line id reporting in verbose, and made T/F RegEx substitution case insensitive
*   Oct 29th, 2016	0.11	fix outline level identification issue
*   Oct 27th, 2016	0.10	initial release on bitbucket
