<h1>Extract PDF to Excel</h1>

This repository contains the code for extracting textual data from a PDF of Faculty Members of IIT Madras and pasting it in an excel sheet. <br>
The code makes use of PyPDF2 and Openpyxl Libraries.

<h2>PyPDF2</h2>
PyPDF2 is a python library used to recognize and extract text from pdf files.<br>
It is an easy to use library which has been used in the code.<br>
Python Library available at: https://pypi.org/project/PyPDF2/

<h2>Openpyxl</h2>
Openpyxl is a python library used to paste the extracted text into an excel file.<br>
Excel files contain multiple sheets and thus the library allows us to choose which sheet we want to edit.<br>
Python library available at: https://pypi.org/project/openpyxl/

<h3>Input</h3>
The input file consists of a PDF conatining data of Faculty Members of IIT Madras

<h3>Output</h3>
The output is an excel file containing data extracted from the PDF file.

<h4>
A few observations:<br>

1. The document being 656 pages long, there are many exceptions in the document.
2. Data needs to be categorized in sections like title, name, designation, department, email, phone, specialties, homepage, etc.
3. Not all the records in the pdf contain all of the needed data and also some of the orders don’t match to other faculty members data.
4. The library is not 100% accurate and misses some of the data but works well for most of the document.

