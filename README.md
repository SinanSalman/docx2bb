# README #


**docx2bb**

## Overview: ##
Create BlackBoard (bb) test questions (text) import file from a Microsoft Word *.docx document. Supported question types are: True/False, Multiple choice, Matching, Essay, and (simple) Fill in the blank. ExamFormat-Sample.docx shows a sample exam format for use with docx2bb. More detailed description of the question identification logic can be found below. Unicode-to-ASCII replacement rules from 'docx2bb.json' data file can be optionally applied.

## Installing docx2bb: ##
Download docx2bb:

* required file: docx2bb.py
* optional file: docx2bb.json [if you choose to change the default unicode2ascii behavior]

In addition to a working python environment, docx2bb requires python-docx and lxml libraries. 
To install these libraries, follow the below steps:

### OSX and Linux: ###
* pip install python-docx

### Windows platforms: ###
* Download lxml library wheel from http://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml
* Install lxml library: pip install [downloadedfile]
* Install python-docx library: pip install python-docx

## Syntax: ##
		python docx2bb.py [options] [docx_filename]
options:	

--verbose	|	-v		display verbose messages

--help		|	-h		display help message

## Docx Formatting and Question Identification Logic: ##
docx2bb requires the use of a simple word format in all questions types to be recognized; specifically, all questions must use an **OUTLINE NUMBERED LIST** format, where questions are listed using level 1 outline and answers use level 2 outline (MAT uses level 3); Any unnumbered paragraph will be ignored by the tool. Key answers for M/C and T/F questions must be selected using **bold** font. 

Essay example:

	1. Question

		a. Answer

The question identification logic is as follows:

* if question has no sub-bullets; it is identified as True/False
* if question has only one sub-bullet; it is identified as Essay
* if question has multiple sub-bullets, but only one is in bold; it is identified as Multiple choice. If a blank is needed in the question use (4) consecutive ‘_’ characters (to avoid being identified as FIB)
* if question has multiple sub-bullets, split evenly between second-level and third-level outline, and none of which are bold; it is identified as Matching
* if question includes (5) or more consecutive ‘_’ characters and no bold answers; it is identified as Fill_in_the_Blank

## Version and History ##
docx2bb version and version history are included in docx2bb.py file header info.

## Disclaimer: ##
docx2bb is provided with no warranties, use it if you find it useful. docx2bb is designed to keep your *.docx document unchanged, but the author assumes no liabilities from use of this tool, including if it eats your exam :).

Code by Sinan Salman, 2016