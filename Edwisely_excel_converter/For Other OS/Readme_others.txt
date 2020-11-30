*****EDWISELY EXCEL GENERATOR (BETA)***** 27/11/20





HOW TO USE:-

1)Install all dependencies

	a)pandas 
		pip install pandas	
	b)docx2txt
		pip install docx2txt
	c)regex
		pip install regex
	d)docx 
		pip install python-docx
	e)pypandoc
		pip install pypandoc
	f)tkinter
	g)os


2)Run "converter.py".GUI should open up.

3)Click on "Select Folder" button.

4)Select the folder where you have stored all the docx files which need conversion.

5)Hit "Convert".

6)Watch magic happen :)


WORKING:-
-->Output will be generated in the same folder which contains docx files.

-->Whenever the converter faces any issue while extracting any latex equation the question is skipped. In such a case
	along with the excel output, an error file is created containing the question numbers 
	which were omitted and the source latex file is generated as well to copy the latex manually.

	KNOWN ISSUES:-
		-While processing matrices.
		(Pls note that these are not all the issues and there may be more issues.)

	 output file names:-
		excel:- "<collegename>_<subjectname>.xlsx"
		error csv:- "errors_<collegename>_<subjectname>.csv"
		source latex:- "<collegename>_<subjectname>.tex"


(Pls note that this application is currently in beta and may throw errors.Feedback in any form is appreciated)

for any queries email me at:- chinmay@edwisely.com



created with love by Chinmay :D