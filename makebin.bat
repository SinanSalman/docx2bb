@Echo off
del docx2bb.pyc 
del docx2bb.spec 
rmdir BuildWIN
rmdir __pycache__

pyinstaller --workpath=.\BuildWIN --distpath=. --onefile .\docx2bb.py
