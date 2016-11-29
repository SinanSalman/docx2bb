#! /bin/bash

rm -R docx2bb.pyc docx2bb.spec BuildOSX
pyinstaller --workpath=./BuildOSX --distpath=. --onefile ./docx2bb.py
