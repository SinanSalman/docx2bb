# -*- coding: utf-8 -*-
"""
docx2bb:
Create BlackBoard (bb) test questions (text) import file from a Microsoft Word *.docx document.
Supported question types are: True/False, Multiple choice, Matching, Essay, and (simple) Fill
in the blank. ExamFormat-Sample.docx shows a sample exam format for use with docx2bb.
Unicode-to-ASCII replacement rules from 'docx2bb.json' data file can be optionally applied.

Licensed under GPLv3
Code by Sinan Salman, 2016-2017
sinan[dot]salman[at]gmail[dot]com
"""

import re
import os
import sys
try:
	import docx2bb_web.mylog as mylog
except:
	import mylog
import json

### Initialization #######################################################################
# encoding=utf8  ==> fix for Python 2.7
if sys.version_info[0] == 2:
	reload(sys)
	sys.setdefaultencoding('utf8')

log = mylog.LOG()
unicode2ascii = {'rules':{'“':'"','”':'"','‘':"'",'’':"'",'–':'-','…':'...','\t':'   '},
				 'notallowed':'[^a-zA-Z0-9 §±!@#$%^&*()\\-_=+[\\]{};:\'\"\\\\|<>,./?`~\\n]'}
data = []
BBtext = ""
QuestionTypes = {'T/F':0, 'M/C':0,'MAT':0,'FIB':0,'ESSAY':0,'Warning':0}

### Analyze Document and Convert to BB Text Format #######################################
def Convert(docx, logfilename='', id=0, ip='0.0.0.0'):
	log.clear('all')
	log.debug('Session ID: {:} ({:})'.format(id,ip))
	try:
		ProcessDocx(docx)
	except ImportError as e:
		log.info('ERROR - ' + e)
	log.save('debug',logfilename)
	return {'result': BBtext,
			'info': log.logtext['info'],
			'debug': log.logtext['debug'],
			'summary': QuestionTypes}

def ProcessDocx(docx):
	"""Process docx contents and create BB text import file"""
	if os.path.isfile('docx2bb.json'):
		if sys.version_info[0] == 2:
			jsonfile = open('docx2bb.json')
		else:
			jsonfile = open('docx2bb.json',encoding="utf8")
		unicode2ascii = json.load(jsonfile)
		log.debug('loaded unicode2ascii from docx2bb.json')

	# extract data from docx file
	global Data
	n=0
	log.debug("Found {:} paragraphs, parsing and converting unicode to ascii...".format(len(docx.paragraphs)))
	for p in docx.paragraphs:
		n += 1
		temp={'No':0,'text':u2a(p.text),'outline':0,'allBold':False,'trueBold':False,'falseBold':False,'list':False}
		bold = True
		for r in p.runs:
			if r.bold == None: bold = False
			if r.bold == True and r.text.strip(' ').lower() == 'true': 	temp['trueBold'] = True
			if r.bold == True and r.text.strip(' ').lower() == 'false': temp['falseBold'] = True
		temp['allBold'] = bold
		temp['list'], temp['outline'] = GetListOutline(p)
		temp['Q'] = 0
		temp['No'] = n
		data.append(temp)

	# log data object for debugging
	log.debug('Before clean up:')
	log.debug('    out   is')
	log.debug(' #  line  list text')
	log.debug('~~~ ~~~~ ~~~~~ ~~~~')
	i = 1
	for d in data:
		log.debug("{:3} {:4} {:5} {:}".format(d['No'],d['outline'],str(d['list']),d['text']))
		i += 1

	log.debug('Clean up:')
	# delete empty paragraphs
	empty_p = []
	for i in range(len(data)):
		if data[i]['text'].strip(' ') == '':
			empty_p.append(i)
	if empty_p != []:
		log.debug('Removing {:} empty line(s): {:}'.format(len(empty_p),[data[x]['No'] for x in empty_p]))
		for i in reversed(range(len(empty_p))): # go backward to delete lines without missing up the index
			del data[empty_p[i]]

	# delete pagebreaks
	PageBreak = []
	for i in range(len(data)):
		if re.search('^\s*\n+$',data[i]['text']):
			PageBreak.append(i)
	if PageBreak != []:
		log.debug('Removing {:} pagebreak(s) at: {:}'.format(len(PageBreak),[data[x]['No'] for x in PageBreak]))
		for i in reversed(range(len(PageBreak))): # go backward to delete lines without missing up the index
			del data[PageBreak[i]]

	# delete non-list paragraphs
	nonlist_p = []
	for i in range(len(data)):
		if data[i]['list'] == False:
			nonlist_p.append(i)
	if nonlist_p != []:
		log.debug('Removing {:} none-list line(s): {:}'.format(len(nonlist_p),[data[x]['No'] for x in nonlist_p]))
		for i in reversed(range(len(nonlist_p))): # go backward to delete lines without missing up the index
			del data[nonlist_p[i]]

	# identify question start/end paragraph positions
	Qbeg_pos = [0]
	Qend_pos = []
	Qid = 1
	data[0]['Q'] = Qid
	for i in range(1, len(data)):
		if data[Qbeg_pos[-1]]['outline'] < data[i]['outline']:
			data[i]['Q'] = Qid
		else:
			Qend_pos.append(i-1)
			Qbeg_pos.append(i)
			Qid += 1
			data[i]['Q'] = Qid
	Qend_pos.append(i)

	# log data object for debugging
	log.debug('After clean up:')
	log.debug('        out   all  true  false is')
	log.debug(' #   Q  line  Bold Bold  Bold  list  text')
	log.debug('~~~ ~~~ ~~~~ ~~~~~ ~~~~~ ~~~~~ ~~~~~ ~~~~')
	for d in data:
		log.debug("{:3} {:3} {:4} {:5} {:5} {:5} {:5} {:}".format(d['No'],d['Q'],d['outline'],str(d['allBold']),str(d['trueBold']),str(d['falseBold']),str(d['list']),d['text']))
	if len(Qbeg_pos) != len(Qend_pos):
		log.info("Error - problem in parsing question lines.")
		log.info("\tQ_begin_positions:\t{:}".format(Qbeg_pos))
		log.info("\t  Q_end_positions:\t{:}".format(Qend_pos))
		return

	# convert to Blackboard import file format
	log.debug('Found {:} possible question(s), identifying type...'.format(len(Qbeg_pos)))
	for i in range(len(Qbeg_pos)):
		make_Q(i, Qbeg_pos[i], Qend_pos[i])

	# prep summary
	log.info('Summary:')
	SumQ = 0
	for k in sorted(QuestionTypes):
		if k != 'Warning':
			log.info('\t{:8}: {:3}'.format(k,QuestionTypes[k]))
			SumQ += QuestionTypes[k]
	log.info('\t~~~~~~~~~~~~~')
	log.info('\t{:8}: {:3}'.format('Total',SumQ))
	if QuestionTypes['Warning'] >0:
		log.info('\t{:8}: {:3}'.format('Warning',QuestionTypes['Warning']))
	log.info('\n')

def GetListOutline(p):
	"""get if paragraph is a list adnd if so its outline level"""

	lst = False
	lvl = 0
	if hasattr(p.paragraph_format.element.pPr, 'numPr'):
		if hasattr(p.paragraph_format.element.pPr.numPr, 'ilvl'):
			lst = True
			if p.paragraph_format.element.pPr.numPr.ilvl != None:
				lvl = p.paragraph_format.element.pPr.numPr.ilvl.val + 1
			else:
				lvl = 1
	if lst == False and lvl == 0:
		if hasattr(p.style.paragraph_format.element.pPr, 'numPr'):
			if hasattr(p.style.paragraph_format.element.pPr.numPr, 'ilvl'):
				lst = True
				if p.style.paragraph_format.element.pPr.numPr.ilvl != None:
					lvl = p.paragraph_format.element.pPr.numPr.ilvl.val + 1
				else:
					lvl = 1
	return lst, lvl

def make_Q(Qid, start, end):
	"""Convert data to Blackboard import file format"""

	global BBtext
	global QuestionTypes

	Qid += 1

	BoldCount = 0	#count number of bold answers. 1 = M/C, 2+ Error
	MAT_start = 0	#Matching answer start position
	if end != start:
		for i in range(start+1,end+1):	# range does not include the end value, so +1 is needed
			if data[i]['allBold']:
				BoldCount += 1
			if MAT_start == 0 and data[i]['outline'] > data[start+1]['outline']:
				MAT_start = i

	# T/F question
	if start == end:
		if data[start]['trueBold'] and data[start]['falseBold']:
			log.info("Warning - skipped T/F question with both answeres in bold. (Q#{:})".format(Qid))
			log.info("\t{:}".format(data[start]['text']))
			QuestionTypes['Warning'] += 1
		elif data[start]['trueBold']:
			Qtxt = re.sub('\([ ]*True[ ]*/[ ]*False[ ]*\)','',data[start]['text'],flags=re.IGNORECASE).strip(' ')
			BBtext += "\nTF\t{:}\ttrue".format(Qtxt)
			log.debug('\tQ{:} identified as True/False'.format(Qid))
			QuestionTypes['T/F'] += 1
		elif data[start]['falseBold']:
			Qtxt = re.sub('\([ ]*True[ ]*/[ ]*False[ ]*\)','',data[start]['text'],flags=re.IGNORECASE).strip(' ')
			BBtext += "\nTF\t{:}\tfalse".format(Qtxt)
			log.debug('\tQ{:} identified as True/False'.format(Qid))
			QuestionTypes['T/F'] += 1
		else:
			log.info("Warning - skipped T/F question with no answeres in bold. (Q#{:})".format(Qid))
			log.info("\t{:}".format(data[start]['text']))
			QuestionTypes['Warning'] += 1
		return

	# Fill In the Blank question
	if BoldCount == 0 and re.search('_{5,}',data[start]['text']) != None:
		BBtext += "\nFIB\t{:}".format(data[start]['text'])
		for i in range(start+1,end+1): # range does not include the end value, so +1 is needed
			BBtext += "\t{:}".format(data[i]['text'])
		log.debug('\tQ{:} identified as Fill_In_the_Blank'.format(Qid))
		QuestionTypes['FIB'] += 1
		return

	# Essay question
	if BoldCount == 0 and start+1 == end:
			BBtext += "\nESS\t{:}\t{:}".format(data[start]['text'],data[end]['text'])
			log.debug('\tQ{:} identified as Essay'.format(Qid))
			QuestionTypes['ESSAY'] += 1
			return

	# M/C question
	if BoldCount == 1:
		BBtext += "\nMC\t{:}".format(data[start]['text'])
		for i in range(start+1,end+1): # range does not include the end value, so +1 is needed
			if data[i]['allBold']: 	answer = 'correct'
			else: 					answer = 'incorrect'
			BBtext += "\t{:}\t{:}".format(data[i]['text'],answer)
		log.debug('\tQ{:} identified as Multiple Choice'.format(Qid))
		QuestionTypes['M/C'] += 1
		return

	# Matching question
	n=0
	if BoldCount == 0 and MAT_start != 0:
		if (end - start)%2 == 0 and MAT_start - start - 1 == end - MAT_start +1 : # equal number of sentences and terms
			BBtext += "\nMAT\t{:}".format(data[start]['text'])
			for i in range(start+1,MAT_start):
				BBtext += "\t{:}\t{:}".format(data[start+1+n]['text'],data[MAT_start+n]['text'])
				n += 1
			log.debug('\tQ{:} identified as Matching'.format(Qid))
			QuestionTypes['MAT'] += 1
		else:
			log.info("Warning - skipped matching question with unequal count of sentances and terms. (Q#{:})".format(Qid))
			log.info("\tterms:{:}, sentances:{:}".format(MAT_start-start-1,end-MAT_start+1))
			log.info("\t{:}".format(data[start]['text']))
			QuestionTypes['Warning'] += 1
		return

	log.info("couldn't identify question type, skipping: (Q#{:})".format(Qid))
	log.info("\t{:}".format(data[start]['text']))
	QuestionTypes['Warning'] += 1

def u2a(txt):
	"""Convert unicode text to ascii"""

	val = txt
	printed = False
	for k,v in unicode2ascii["rules"].items():
		if txt.find(k) != -1:
			if not printed:
				printed = True
				log.debug("\t{:}".format(txt))
			log.debug("\t\tconverted {:} to {:}".format(k,v))
			val = val.replace(k,v)
	found_itr = re.finditer(unicode2ascii["notallowed"],val)
	found_pos = [m.start()+1 for m in found_itr]
	found_val = [txt[m-1] for m in found_pos]
	if found_pos != []:
		if not printed:
			log.debug("\t{:}".format(txt))
		log.debug("found {:} unhandled unicode at position(s) {:}".format(found_val,found_pos))
	return val
