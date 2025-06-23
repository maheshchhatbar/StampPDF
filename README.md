# StampPDF
Print a PDF on a Letterhead using stamp feature of pdftk.exe

StampPDF is a simple utility I made for myself to Print/Stamp softcopies oof Invoices on to sofcopies of Letterheads. Generally the softcopies of the invoices were suitable printing as hardcopies to be printed on to physical letterheads. But since some time I needed to submit the softcopies which were visibly on letterheads. I came across the utility pdftk.exe which did serve my purpose , the free server uses command line interface. So as to overcome typing hassle and use drag and drop feature this project was made. 
The textboxes used are for : 

1 - the path of pdftk.exe , the default install location generally used is mentioned , if pdftk is not found the utility will try to search for pdftk from the path , if path to pdftk.exe is found the textbox will be disabled , or you can specify it if needed.

2 - the letterhead as a pdf on to which the other pdfs are to be printed/stamped , pls paste the path and filename of the letterhead to be used , if found the textbox will be disabled.

3 - Prepend to filename can be set as required to suit your needs so as to avoid overwriting the existing pdf , generally files printed on letterheads prepend filenames with "LH - " which is set as default but you can cahnge it to whatever you prefer.

4 - Append to filename can be set as required to suit your needs so as to avoid overwriting the existing pdf.

5 - actual shell command executed , it can be copy pasted into a cmd and it should give the same expected result

Useage : Drag and Drop the pdf that you want to stamp on this form and a new pdf will be generated and will be saved in the same directory of the pdf that has to be stamped with the "Prepend" and "Append" added in the filename. The file name of the new file generated will be shown in the last textbox.

The links to all the source codes used and from where the respective codes are taken/used is included in respective modules for anyone to see.
I have made and tested this utility with the pdftk.exe --version 2.02 available from 

https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/pdftk_server-2.02-win-setup.exe 

This is a VB6 application/project.

I hope that this utility is useful to anyone.
Many improvements are needed , this is just a start for anyone to use this utility.

Please do let me know if you feel I can improve the utility.

