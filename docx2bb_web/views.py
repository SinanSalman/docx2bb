import os
import json
import time
import shutil
import atexit
import pickle
import docx
from docx2bb_web import app
import docx2bb_web.docx2bb_lib as d2b
from flask import request, session, redirect, url_for, render_template, flash, Response

# flask setup
app.config.from_object(__name__)  # load config from this file
with open('./docx2bb_web/config.json') as config_file:  # Load default config and override config from an environment variable
	config_data = json.load(config_file)
config_data['LOGFILE'] = os.path.join(app.instance_path, config_data['LOGFILE'])
config_data['STATEFILE'] = os.path.join(app.instance_path, config_data['STATEFILE'])
app.config.update(config_data)
app.config.from_envvar('docx2bb_SETTINGS', silent=True)
app.secret_key = 'change to a random value and keep this really secret'  # set the secret key for 'session'
NextID = 1000
ALLOWED_EXTENSIONS = set(['docx'])


################################################################################
# utility functions
################################################################################
@atexit.register  # this won't work in Flask-debug-mode as the restart process will change the variable ID
def cleanup():
	if save_state():
		print('\nSaved state. Server shutting down...')
	else:
		print('\nFailed to save state. Server shutting down...')


def save_state():
	with open(app.config['STATEFILE'],'wb') as f:
		pickle.dump((NextID,app.secret_key),f)
		return True
	return False


def load_state():
	if os.path.isfile(app.config['STATEFILE']):
		with open(app.config['STATEFILE'],'rb') as f:
			NextID,app.secret_key = pickle.load(f)
			return True
	return False


# loading previous server state data if existing
load_state()


def allowed_file(filename):
	return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def make_new_secret_key():
	app.secret_key = os.urandom(24)


################################################################################
# application views
################################################################################
@app.route('/shutdown', methods=['GET'])
def shutdown():
	func = request.environ.get('werkzeug.server.shutdown')
	if func is None:
		raise RuntimeError('Not running with the Werkzeug Server')
	func()
	return 'Server shutting down...'


@app.route('/reset_secret', methods=['GET'])
def reset_secret():
	make_new_secret_key()
	return redirect(url_for('index'))


@app.route('/')
@app.route('/index')
def index():
	global NextID
	if 'ID' not in session:
		session['ID'] = NextID
		NextID += 1
	if 'output' not in session:
		session['output'] = {}
	if 'timestamp' in session['output']:
		if time.time() - session['output']['timestamp'] > app.config['RESULTS_LIFE_TIME']:
			session['output'] = {}
	session['output']['lifetime'] = app.config['RESULTS_LIFE_TIME']
	return render_template('index.html',output=session['output'])


@app.route('/about')
def about():
	return render_template('about.html')


@app.route('/tutorial')
def tutorial():
	return render_template('tutorial.html')


@app.route('/logout')
def logout():
	session.pop('username',None)
	session.pop('password',None)
	return redirect(url_for('index'))


@app.route('/admin')
@app.route('/admin_login')
def admin_login():
	if 'password' in session.keys() and 'username' in session.keys():
		if app.config['USERNAME'] == session['username'] and app.config['PASSWORD'] == session['password']:
			return redirect(url_for('admin_server'))
	return render_template('admin_login.html')


@app.route('/admin_server', methods=['POST','GET'])
def admin_server():
	if request.method == 'GET':
		if 'password' in session.keys() and 'username' in session.keys():
			if app.config['USERNAME'] != session['username']:
				flash('User is not an admin!')
				return redirect(url_for('admin_login'))
			if app.config['PASSWORD'] != session['password']:
				flash('Incorrect password!')
				return redirect(url_for('admin_login'))
		else:
			return redirect(url_for('admin_login'))
	elif request.method == 'POST':
		if app.config['USERNAME'] != request.form['username']:
			flash('User is not an admin!')
			return redirect(url_for('admin_login'))
		if app.config['PASSWORD'] != request.form['password']:
			flash('Incorrect password!')
			return redirect(url_for('admin_login'))
		session['username'] = request.form['username']
		session['password'] = request.form['password']
	if os.path.isfile(app.config['LOGFILE']):
		with open(app.config['LOGFILE'],'r') as f:
			log = f.read()
	else:
		log = ''
	return render_template('admin_server.html',log=log)


@app.route('/start_new_log', methods=['GET'])
def start_new_log():
	if 'password' in session.keys() and 'username' in session.keys():
		if app.config['USERNAME'] != session['username']:
			flash('User is not an admin!')
			return redirect(url_for('admin_login'))
		if app.config['PASSWORD'] != session['password']:
			flash('Incorrect password!')
			return redirect(url_for('admin_login'))
	else:
		return redirect(url_for('admin_login'))
	if os.path.isfile(app.config['LOGFILE']):
		shutil.move(app.config['LOGFILE'],app.config['LOGFILE']+'.'+time.strftime("%m-%d-%Y %H:%M:%S"))
	return redirect(url_for('admin_server'))


@app.route('/load_docx', methods=['POST'])
def load_docx():
	if 'filename' not in request.files:
		flash('No file part in posted request.')
		return redirect(url_for('index'))
	file = request.files['filename']
	if file.filename == '':
		flash('No selected file.')
		return redirect(url_for('index'))
	if file and allowed_file(file.filename):
		session['output'] = d2b.Convert(docx.Document(file),logfilename=app.config['LOGFILE'],id=session['ID'], ip=request.remote_addr)
		session['output']['filename'] = file.filename.replace('.docx','.txt').replace('.doc','.txt')
		session['output']['timestamp'] = time.time()
	else:
		flash('Can\'t load file.')
	return redirect(url_for('index'))


@app.route('/download_txt', methods=['GET'])
def download_txt():
	try:
		if time.time() - session['output']['timestamp'] > app.config['RESULTS_LIFE_TIME']:
			flash("Converted file is older than {:} seconds. All data were purged, please convert it again and downlod the results within the time allowed.".format(app.config['RESULTS_LIFE_TIME']))
			return redirect(url_for('index'))
		return Response(session['output']['result'].strip(' ').strip('\n'), mimetype="text/plain", headers={"Content-disposition":"attachment; filename=" + session['output']['filename']})
	except Exception as e:
		flash("Error while getting converted file. " + str(e))
		return redirect(url_for('index'))
