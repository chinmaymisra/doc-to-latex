*****EDWISELY EXCEL GENERATOR (BETA)***** 27/11/20





HOW TO USE:-

NOTE:- Method 1 is a bit complex but foolproof and while Method 2 is less tedious it may not work on some systems.

METHOD-1:

1)Install Python,pls follow online instructions, don't forget to add to path.
(Link for adding python to path :- https://www.educative.io/edpresso/how-to-add-python-to-path-variable-in-windows, create a new environment variable
called Pythonpath with the same value as well.)

(Python 3.7.9 was used when creating this application, no errors have been encountered with newer builds of python but in case of encountering
some module incompatibility issue first try with a this older build of python before reporting the issue)

https://www.python.org/ftp/python/3.9.0/python-3.9.0-amd64.exe(for 64 bit machines)

2)Press Windows+R to open CMD 


3)Install the following modules using Pip(pip is a python package installer):-
	
	use command:"pip install <package name>==<version>"
	
	a)tkinter(8.6.8)
	b)pandas(0.25.1)
	c)docx2txt(0.8)
	d)regex(2020.10.28)
	e)docx (0.2.4)
	f)pypandoc(1.5)
	g)pandoc(2.0a5)

	IMP:-
	pls run the below command at the end,current pandas installation installs numpy(1.19.4) by default which causes compatibility issues with pypandoc.
	"pip install numpy==1.19.3"

4)Change directory to the directory of the "converter.py" file using "cd" command

5)Run "python converter.py" in CMD.Application should open.

6)Click on "Select Folder" button.

7)Select the folder where you have stored all the docx files which need conversion.

8)Hit "Convert".

9)Watch magic happen :)
				
					OR

METHOD-2:

1)Go to "dist" folder.

2)Click on "converter.exe".

3)Click on "Select Folder" button.

4)Select the folder where you have stored all the docx files which need conversion.

5)Hit "Convert".

6)Watch magic happen :)




NOTE:- 	There is a specific format for the input docx files,
	a template and some examples have been added in the "Examples" folder.
	You can try out with those first to ensure the script is working.
	Please create further docx files according the template format.


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