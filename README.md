# README #

docx2bb

# Overview: #
Create BlackBoard (bb) test questions (text) import file from a MS Word *.docx document.
Supported question types: True/False, Multiple choice, Matching, Essay, and Fill in the blank. 
ExamFormat-Sample.docx shows a sample exam format for use with docx2bb. More detailed 
description of the question identification logic is listed below. Unicode-to-ASCII replacement 
rules from 'docx2bb.json' data file can be applied optionally.

# Installing doc2bb: #
In addition to a working python environment, docx2bb requires python-docx and lxml libraries. 
To install these libraries, follow the below steps:

OSX and Linux:
	pip install python-docx

Windows platforms:
	Download lxml library wheel:
		http://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml
	Install lxml library:
		pip install [downloadedfile]
	Install python-docx library:
		pip install python-docx

# Syntax: #
		python docx2bb.py [options] [docx_filename]
options:
		--verbose	|	-v		display verbose messages
		--help		|	-h		display help message

# Docx Formatting and Question Identification Logic: #
docx2bb requires the use of a simple word format in all questions types to be recognized; 
specifically, all questions must use an OUTLINE NUMBERED LIST format, where questions are
listed using level 1 outline and answers use level 2 outline; Any unnumbered paragraph will 
be ignored by the tool. Key answers for questions (except for Essay and Fill_in_the_Blnak) 
must be selected using bold font. 

Example:
	1. Question
		a. Answer

The question identification logic is as followes:
- if question includes five consecutive ‘_’ characters; it is identified as Fill_in_the_Blank
- if question does not have any sub-bullets; it is identified as True/False
- if question has only one sub-bullet; it is identified as Essay
- if question has multiple sub-bullets, but only one is in bold; it is identified as Multiple choice
- if question has multiple sub-bullets, and more than one are in bold; it is identified as Matching

# Disclaimer: #
docx2bb is provided with no warranties, use it if you find it useful. docx2bb was designed
to keep your *.docx exam unchanged, but the author assumes no liabilities from use or 
misuse of this tool.

Code by Sinan Salman, 2016