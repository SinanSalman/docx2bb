#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
docx2bb

* Overview:
Create BlackBoard (bb) test questions (text) import file from a MS Word *.docx document.
Supported question types: True/False, Multiple choice, Matching, Essay, and Fill in the blank. 
ExamFormat-Sample.docx shows a sample exam format for use with docx2bb. More detailed 
description of the question identification logic is listed below. Unicode-to-ASCII replacement 
rules from 'docx2bb.json' data file can be applied optionally.

* Installing doc2bb:
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

* Syntax:
		python docx2bb.py [options] [docx_filename]
options:
		--verbose	|	-v		display verbose messages
		--help		|	-h		display help message

* Docx Formatting and Question Identification Logic:
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

* Disclaimer:
docx2bb is provided with no warranties, use it if you find it useful. docx2bb was designed
to keep your *.docx exam unchanged, but the author assumes no liabilities from use or 
misuse of this tool.

Code by Sinan Salman, 2016

Version History:
21.10.27	0.10	Initial release

"""

HELP_MSG = """Options:
	--verbose | -v  display verbose messages
	--help    | -h  display help message
"""

INSTALL_MSG = """
OSX and Linux:
	pip install python-docx

Windows platforms:
	Download lxml library wheel:
		http://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml
	Install lxml library:
		pip install [downloadedfile]
	Install python-docx library:
		pip install python-docx
"""
		
__app__ = 		"docx2bb.py"
__author__ = 	"Sinan Salman (sinan.salman[at]gmail.com)"
__version__ = 	"v0.10"
__date__ = 		"Oct 27, 2016"
__copyright__ = "Copyright (c)2016 Sinan Salman"
__license__ = 	"GPLv3"
__website__	=	"https://bitbucket.org/sinansalman/docx2bb"

### Initialization #######################################################################

import re
import os
import sys
import platform
        
# encoding=utf8  ==> fix for Python 2.7
if sys.version_info[0] == 2:
	reload(sys)  
	sys.setdefaultencoding('utf8')

verbose = False
unicode2ascii = {'rules':None,'notallowed':None}
WordFileName = ""
data = []
BBtext = ""
TextWidth = 0
QuestionTypes = {'T/F':0, 'M/C':0,'MAT':0,'FIB':0,'ESSAY':0,'Warning':0}

if platform.system() == 'Windows':
        class bcolors:
                WARNING = 	''
                FAIL = 		''
                ENDC = 		''
else:
        class bcolors:
                WARNING = 	'\033[33m'
                FAIL = 		'\033[31m'
                ENDC = 		'\033[0m'

### Script Management ####################################################################

def ProcessCLI():
	"""Process CLI parameters"""
	global verbose
	global WordFileName
	global unicode2ascii
	global TextWidth
	
	try:
		TextWidth = os.get_terminal_size()[0]
	except:
		TextWidth = 80

	print ('{:} | {:} | {:} | {:} License\nDownload latest version at {:}'.format(__app__,__version__,__date__,__license__,__website__))
	
	# handle arguments and load settings from JSON file
	if len(sys.argv) == 1:
		print ("\nSyntax:\n\tpython docx2bb.py [options] [docx_filename]")
		print (HELP_MSG)
		print_Fail ("Missing argument")
	if '--verbose' in sys.argv or '-v' in sys.argv:
		print ("\n*** Option: verbose mode")
		verbose = True
	if '--help' in sys.argv or '-h' in sys.argv:
		print("\nSyntax:\n\tpython docx2bb.py [options] [docx_filename]")
		print(HELP_MSG)
		sys.exit(0)
	WordFileName = sys.argv[-1]
	if not os.path.isfile(WordFileName):
		print_Fail ("can't find file: {:}. Make sure [docx_filename] is the last argument.".format(WordFileName))
	import json
	if os.path.isfile('docx2bb.json'):
		with open('docx2bb.json',encoding='utf-8') as jsonfile:
			unicode2ascii = json.load(jsonfile)
	
### Analyze Document and Convert to BB Text Format #######################################

def RunScript():
	"""Process word file and create BB text import file"""

	# try to import python-docx module
	try:
		import docx
	except ImportError as e:
		print ("Please install python-docx library first using:")
		print (INSTALL_MSG)
		sys.exit(0)

	if verbose: print ('Reading Docx file...')
	doc = docx.Document(WordFileName)

	# extract data from docx file
	global Data
	if verbose: print ("Found {:} paragraphs, parsing and converting unicode to ascii...".format(len(doc.paragraphs)))
	for p in doc.paragraphs:
		temp={'text':u2a(p.text),'left_indent':0,'allBold':False,'trueBold':False,'falseBold':False,'list':False}
		if p.paragraph_format.left_indent != None:
			temp['left_indent'] = p.paragraph_format.left_indent
		elif p.style.paragraph_format.left_indent != None:
			temp['left_indent'] = p.style.paragraph_format.left_indent
		bold = True
		for r in p.runs:
			if r.bold == None: bold = False
			if r.bold == True and r.text.strip(' ').lower() == 'true': 	temp['trueBold'] = True
			if r.bold == True and r.text.strip(' ').lower() == 'false': temp['falseBold'] = True
		temp['allBold'] = bold
		temp['list'] = isList(p.style)
		temp['Q'] = 0
		data.append(temp)

	# create an outline level field
	indent_seq = list(set([x['left_indent'] for x in data]))
	indent_seq.sort()
	for i in range(len(data)):
		data[i]['outline'] = indent_seq.index(data[i]['left_indent'])

	if verbose: print_data_table('Before clean up:') # print data object for debugging

	if verbose:	print ('\nClean up:')	
	# delete empty paragraphs
	empty_p = []
	for i in range(len(data)):
		if data[i]['text'].strip(' ') == '':
			empty_p.append(i)
	if empty_p != []:
		if verbose:
			print ('Removing {:} empty line(s): {:}'.format(len(empty_p),[x+1 for x in empty_p]))	
		for i in reversed(range(len(empty_p))): # go backward to delete lines without missing up the index
			del data[empty_p[i]]

	# delete non-list paragraphs
	nonlist_p = []
	for i in range(len(data)):
		if data[i]['list'] == False:
			nonlist_p.append(i)
	if nonlist_p != []:
		if verbose:
			print ('Removing {:} none-list line(s): {:}'.format(len(nonlist_p),[x+1 for x in nonlist_p]))	
		for i in reversed(range(len(nonlist_p))): # go backward to delete lines without missing up the index
			del data[nonlist_p[i]]

	# identify question start/end paragraph positions
	Qbeg_pos = [0]
	Qend_pos = []
	Qid = 1
	data[0]['Q'] = Qid
	for i in range(1, len(data)-1):
		if data[Qbeg_pos[-1]]['outline'] < data[i]['outline']:
			data[i]['Q'] = Qid
		else:
			Qend_pos.append(i-1)
			Qbeg_pos.append(i)
			Qid += 1
			data[i]['Q'] = Qid
	Qend_pos.append(len(data)-1) 
	data[-1]['Q'] = Qid

	if verbose: print_data_table('After clean up:') # print data object for debugging
	if len(Qbeg_pos) != len(Qend_pos): print_Fail ("problem in parsing question lines.\n\tQ_begin_positions:\t{:}\n\tQ_end_positions:\t{:}".format(Qbeg_pos,Qend_pos))

	# convert to Blackboard import file format
	if verbose:
		print ('\nFound {:} possible question(s), identifying type...'.format(len(Qbeg_pos)))
	for i in range(len(Qbeg_pos)):
		make_Q(i, Qbeg_pos[i], Qend_pos[i])

	# write to Blackboard text file
	with open(WordFileName.replace('.docx','.txt'),'w') as outputfile:
		outputfile.write(BBtext.strip('\n'))
		outputfile.close()
	
	print ('\nSuccessfully wrote question(s) to [{:}]. Summary:'.format(WordFileName.replace('.docx','.txt')))
	SumQ = 0
	for k in sorted(QuestionTypes):
		if k != 'Warning':
			print ('\t{:8}: {:3}'.format(k,QuestionTypes[k]))
			SumQ += QuestionTypes[k]
	print ('\t~~~~~~~~~~~~~\n\t{:8}: {:3}'.format('Total',SumQ))
	if QuestionTypes['Warning'] >0:
		print ('\t{:8}: {:3}'.format('Warning',QuestionTypes['Warning']))
	
def make_Q(Qid, start, end):
	"""Convert data to Blackboard import file format"""

	global BBtext
	global QuestionTypes
	
	Qid += 1

	if re.search('_{5,}',data[start]['text']) != None:	# Fill In the Blank question
		BBtext += "\nFIB\t{:}".format(data[start]['text'])
		for i in range(start+1,end+1): # range does not include the end value, so +1 is needed
			BBtext += "\t{:}".format(data[i]['text'])
		if verbose: print ('\tQ{:} identified as Fill_In_the_Blank'.format(Qid))	
		QuestionTypes['FIB'] += 1
		return

	if start == end:	# T/F question
		if data[start]['trueBold'] and data[start]['falseBold']:
			print_Warning ("skipped T/F question with both answeres in bold. (Q#{:})\n\t{:}...".format(Qid,data[start]['text'][:(TextWidth-15)]))
			QuestionTypes['Warning'] += 1
		elif data[start]['trueBold']:
			BBtext += "\nTF\t{:}\ttrue".format(data[start]['text'])
			if verbose: print ('\tQ{:} identified as True/False'.format(Qid))
			QuestionTypes['T/F'] += 1
		elif data[start]['falseBold']:
			BBtext += "\nTF\t{:}\tfalse".format(data[start]['text'])
			if verbose: print ('\tQ{:} identified as True/False'.format(Qid))
			QuestionTypes['T/F'] += 1
		else:
			print_Warning ("skipped T/F question with no answeres in bold. (Q#{:})\n\t{:}...".format(Qid,data[start]['text'][:(TextWidth-15)]))
			QuestionTypes['Warning'] += 1
		return

	if start+1 == end:	# Essay question
			BBtext += "\nESS\t{:}\t{:}".format(data[start]['text'],data[end]['text'])
			if verbose: print ('\tQ{:} identified as Essay'.format(Qid))
			QuestionTypes['ESSAY'] += 1
			return

	BoldCount = 0	#count number of bold answers. 1 = M/C, 2+ MAT
	RegCount = 0	#count number of regular answers
	RegEndPos = -1
	BoldStartPos = 99999
	for i in range(start+1,end+1):	# range does not include the end value, so +1 is needed
		if data[i]['allBold']: 
			BoldCount += 1
			if i < BoldStartPos: BoldStartPos = i
		else:
			RegCount += 1
			if i > RegEndPos: RegEndPos = i
 
	if BoldCount == 1:	# M/C question
		BBtext += "\nMC\t{:}".format(data[start]['text'])
		for i in range(start+1,end+1): # range does not include the end value, so +1 is needed
			if data[i]['allBold']: 	answer = 'correct'
			else: 					answer = 'incorrect'
			BBtext += "\t{:}\t{:}".format(data[i]['text'],answer)
		if verbose: print ('\tQ{:} identified as Multiple Choice'.format(Qid))	
		QuestionTypes['M/C'] += 1
		return

	if BoldCount > 1 and BoldStartPos>RegEndPos:	# Matching question
		if BoldCount == RegCount and end-start == BoldCount+RegCount: # inclusive of end:start (so total count is end-start+1)
			BBtext += "\nMAT\t{:}".format(data[start]['text'])
			for i in range(start+1,start+BoldCount+1): # range does not include the end value, so +1 is needed
				BBtext += "\t{:}\t{:}".format(data[i]['text'],data[i+BoldCount]['text'])
			if verbose: print ('\tQ{:} identified as Matching'.format(Qid))
			QuestionTypes['MAT'] += 1
		else:
			print_Warning ("skipped matching question with unequal count of sentances and terms. (Q#{:})\n\tterms:{:}, sentances:{:}\n\t{:}...".format(Qid,BoldCount,end-start-1-BoldCount,data[start]['text'][:(TextWidth-15)]))
			QuestionTypes['Warning'] += 1
		return

	print_Warning ("couldn't identify question type, skipping: (Q#{:}, best guess: ESS or M/C)\n\t{:}...".format(Qid,data[start]['text'][:(TextWidth-15)]))
	QuestionTypes['Warning'] += 1
	
def u2a(txt):
	"""Convert unicode text to ascii"""
	
	val = txt
	printed = False
	if unicode2ascii["rules"] != None:
		for k,v in unicode2ascii["rules"].items():
			if verbose and txt.find(k) != -1:
				if not printed:
					printed = True
					print_to_console ("\t{:}...".format(txt[:(TextWidth-15)]))
				print_to_console ("\t\tconverted {:} to {:}".format(k,v))
			val = val.replace(k,v)
	if verbose and unicode2ascii["notallowed"] != None:
		found_itr = re.finditer(unicode2ascii["notallowed"],val)
		found_pos = [m.start()+1 for m in found_itr]
		found_val = [txt[m-1] for m in found_pos]
		if found_pos != []:
			if not printed:
				print_to_console ("\t{:}...".format(txt[:(TextWidth-15)]))
			print_Warning ("found {:} unhandled unicode at {:}".format(found_val,found_pos)) 
	return val

def isList(style):
	""" recursevely go through style and base styles to check if it is list paragraph"""
	if style == None:
		return False
	elif str(style).find('List Paragraph') != -1:
		return True
	else:
		return isList(style.base_style)

def	print_data_table(title):
	print ('\n{:}'.format(title))
	print ('        out    left \tall \ttrue\tfalse\tis        ') 
	print (' #   Q  line  indent\tBold\tBold\tBold \tlist\ttext') 
	print ('~~~ ~~~ ~~~~  ~~~~~~\t~~~~\t~~~~\t~~~~~\t~~~~\t~~~~') 
	i = 1
	for d in data:
		print ("{:3} {:3} {:4}  {:6}\t{:}\t{:}\t{:}\t{:}\t{:}".format(i,d['Q'],d['outline'],d['left_indent'],d['allBold'],d['trueBold'],d['falseBold'],d['list'],d['text'][:(TextWidth-56)])) 
		i += 1

### print to console string with any kind of encoding ####################################

def print_Warning(msg):
	"""Print warning message with color"""
	print_to_console (bcolors.WARNING + 'Warning - ' + msg + bcolors.ENDC)

def print_Fail(msg):
	"""Print fail message with color and exit"""
	print_to_console (bcolors.FAIL + 'Error - ' + msg + bcolors.ENDC)
	sys.exit(0)

def print_to_console(text):
	"""Prints a (unicode) string to the console, encoded depending on the stdout encoding 
	(eg. cp437 on Windows). Works with Python 2 and 3."""	
	try:
		sys.stdout.write(text)
	except UnicodeEncodeError:
		bytes_string = text.encode(sys.stdout.encoding, 'backslashreplace')
		if hasattr(sys.stdout, 'buffer'):
			sys.stdout.buffer.write(bytes_string)
		else:
			text = bytes_string.decode(sys.stdout.encoding, 'strict')
			sys.stdout.write(text)
	sys.stdout.write("\n")
    
### Main #################################################################################

if __name__ == "__main__":
	ProcessCLI()
	RunScript()
