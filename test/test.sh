#! /bin/bash

echo "test script for docx2bb.py"
echo $(date) > test.log

echo ""								>> test.log
echo "test options/parameters error messages:" >> test.log
echo ""								>> test.log

echo "$>python docx2bb.py" 			>> test.log
python ../docx2bb.py 					>> test.log
echo "$>python docx2bb.py -h"		>> test.log
python ../docx2bb.py -h  				>> test.log
echo "$>python docx2bb.py -v" 		>> test.log
python ../docx2bb.py -v  				>> test.log
echo "$>python docx2bb.py BadName" 	>> test.log
python ../docx2bb.py BadName  			>> test.log

echo ""								>> test.log
echo "****************************"	>> test.log
echo "****************************"	>> test.log
echo "****************************"	>> test.log
echo ""								>> test.log

echo "test different exam formats:" >> test.log


SAVEIFS=$IFS
IFS=$(echo -en "\n\b")

for f in *.docx
do
	echo ""								>> test.log
	echo "****************************"	>> test.log
	echo "****************************"	>> test.log
	echo "****************************"	>> test.log
	echo ""								>> test.log

	name=$(echo $f | cut -f1 -d'.')
	echo "$>python docxbb.py -v $name.docx" >> test.log
	python ../docx2bb.py "$name.docx"	 	>> test.log
	diff "$name.txt" "$name.BM.txt"			>> test.log
done

IFS=$SAVEIFS

echo "done. check [test.log] file"
