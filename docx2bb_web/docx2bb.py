#!/usr/bin/env python
"""
docx2bb:
Create BlackBoard (bb) test questions (text) import file from a Microsoft Word *.docx document.
Supported question types are: True/False, Multiple choice, Matching, Essay, and (simple) Fill
in the blank. ExamFormat-Sample.docx shows a sample exam format for use with docx2bb.
Unicode-to-ASCII replacement rules from 'docx2bb.json' data file can be optionally applied.

Syntax:
	docx2bb [options] [docx_filename]
or
	python docx2bb.py [options] [docx_filename]
options:
	--verbose  | -v  display verbose messages
	--help     | -h display help message

Disclaimer:
docx2bb is provided with no warranties, use it if you find it useful. docx2bb is designed to
keep your *.docx document unchanged, but the author assumes no liabilities from use of
this tool, including if it eats your exam ;).
"""

HELP_MSG = """Options:
	--verbose | -v  display verbose messages
	--help    | -h  display help message
"""

INSTALL_MSG = """
OSX, Linux, and Windows:
	pip install python-docx

If missing lxml library on Windows platforms:
	Download lxml library wheel:
		http://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml
	Install lxml library:
		pip install [downloadedfile]
	Install python-docx library:
		pip install python-docx
"""

__app__ = "docx2bb"
__author__ = "Sinan Salman (sinan[dot]]salman[at]gmail[dot]com)"
__version__ = "v0.21"
__date__ = "Feb 24, 2019"
__copyright__ = "Copyright (c)2016-2019 Sinan Salman"
__license__ = "GPLv3"
__website__ = "https://bitbucket.org/sinansalman/docx2bb"

import os
import sys
import docx2bb_lib as d2b

# Initialization ###############################################################
verbose = False
WordFileName = ""


# Script Management ############################################################
def ProcessCLI():
	"""Process CLI parameters"""
	global verbose
	global WordFileName

	# Get terminal width
	try:
		d2b.log.TextWidth = os.get_terminal_size()[0]
	except:
		d2b.log.TextWidth = 80

	print('{:} | {:} | {:} | {:} License\nDownload latest version at {:}\n'.format(__app__,__version__,__date__,__license__,__website__))

	# handle arguments
	if len(sys.argv) == 1:
		print("Syntax:\n\tdocx2bb [options] [docx_filename]\n\tpython docx2bb.py [options] [docx_filename]")
		print(HELP_MSG)
		print("Error - Missing argument")
		sys.exit(0)
	if '--verbose' in sys.argv or '-v' in sys.argv:
		print("*** Option: verbose mode")
		verbose = True
	if '--help' in sys.argv or '-h' in sys.argv:
		print("Syntax:\n\tdocx2bb [options] [docx_filename]\n\tpython docx2bb.py [options] [docx_filename]")
		print(HELP_MSG)
		sys.exit(0)
	WordFileName = sys.argv[-1]
	if not os.path.isfile(WordFileName):
		print("Error - can't find file: {:}. Make sure [docx_filename] is the last argument.".format(WordFileName))


# Analyze Document and Convert to BB Text Format ###############################
def RunScript():
	"""Process word file and create BB text import file"""

	# try to import python-docx module
	try:
		import docx
	except ImportError as e:
		print("Error importing library. If using docx2bb.py directly please install [python-docx] first using:")
		print(INSTALL_MSG)
		sys.exit(0)

	# open docx file
	if verbose:
		print('Reading Docx file...\n')
	output = d2b.Convert(docx.Document(WordFileName),'activity_cli.log')

	# write to Blackboard text file
	with open(WordFileName.replace('.docx','.txt'),'w') as outputfile:
		outputfile.write(output['result'].strip('\n'))

	if verbose:
		print_to_console(output['debug'])
	else:
		print_to_console(output['info'])


# print to console string with any kind of encoding ############################
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


# Main #########################################################################
if __name__ == "__main__":
	ProcessCLI()
	RunScript()
