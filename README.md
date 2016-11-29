# README #


**docx2bb**

## Overview: ##
Create BlackBoard (bb) test questions (\*.txt) upload file from a Microsoft Word (\*.docx) test document. Supported question types are: True/False, Multiple choice, Matching, Essay, and (simple) Fill in the blank. **ExamFormat-Sample.docx** provides a sample exam format for use with docx2bb; the critical requirement for the word document is to format all questions and answers as an **outline numbered list**. More detailed description of the question identification logic can be found below. 

## Downloading docx2bb: ##
Choose one of the below described files depending on your computing environment:

* docx2bb     (OSX binary)
* docx2bb.exe (Windows binary)
* docx2bb.py  (python code, see below for the python environment requirements)

## Syntax: ##
		docx2bb [options] [docx_filename]
or
		python docx2bb.py [options] [docx_filename]

options:	

--verbose	|	-v		display verbose messages

--help		|	-h		display help message

## Python Environment Requirements: ##

If you choose to use the python code directly (not the binary), you'll need a working python environment as well as python-docx and lxml libraries. To install these libraries, on top of a python environment follow the below steps:

### OSX and Linux: ###
* pip install python-docx

### Windows platforms: ###
* Download lxml library wheel from http://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml
* Install lxml library: pip install [downloadedfile]
* Install python-docx library: pip install python-docx

## Unicode-to-ASCII replacement rules: ##
Optionally, Unicode-to-ASCII replacement rules can be changed by downloading and modifying the 'docx2bb.json' data file. Default replacement rules will be used if no such file is found in the local folder.

## Docx Formatting and Question Identification Logic: ##
docx2bb requires the use of a simple word format for questions types to be recognized; specifically, all questions must use an **OUTLINE NUMBERED LIST** format, where questions are listed using level 1 outline and answers use level 2 outline (MAT uses level 3); Any unnumbered paragraph will be ignored by the tool. Key answers for M/C and T/F questions must be selected using **bold** font. 

Essay example:

	1. Question

		a. Answer

The question identification logic is as follows:

* if question has no sub-bullets; it is identified as True/False.
* if question has only one sub-bullet; it is identified as Essay.
* if question has multiple sub-bullets, but only one is in bold; it is identified as Multiple choice. If a blank is needed in the question use (4) consecutive '_' characters (to avoid being identified as FIB).
* if question has multiple sub-bullets, split evenly between second-level and third-level outline, and none of which are bold; it is identified as Matching.
* if question includes (5) or more consecutive ‘_’ characters and no bold answers; it is identified as Fill_in_the_Blank. Multiple possible answers are allowed in this type.

## Version and History ##
docx2bb version and version history are included in docx2bb.py file header info.

## License: ##
docx2bb is licensed under GPLv3.0 which can be accessed at https://www.gnu.org/licenses/gpl-3.0.en.html

## Disclaimer: ##
docx2bb is provided with no warranties, use it if you find it useful. docx2bb is designed to keep your *.docx document unchanged, but the author assumes no liabilities from use of this tool, including if it eats your homework/exam :).

Code by Sinan Salman, 2016