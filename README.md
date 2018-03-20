# generate signature letter from excel file

This is a python script to generate a letter of signatures from an excel file. Written specifically for letters in the house of reps, but could be modified for whatever. Note that I think this only works with python2 currently. Sorry.

I have no idea how transportable this will be. Hopefully someone will let me know. 

## running the script

To execute the script, download `letter_sigs_commandline.py`, `input_file.xlsx`, and `custom_styles.docx`.  While in the directory containing the input files, from the command line run:

`python letter_sigs_commandline.py -i input_file.xlsx -s custom_styles.docx -t 'preferred title' -o out_file.docx`

This will generate the docx file `out_file.docx` which will contain the appropriate signatures. Change the names of these files as you see fit.

'preferred title' can be changed to whatever you want. It will be the same for each individual. Make sure to include it in apostrophes. 

## required packages

You'll need the following to run this package:

1. [python2](https://www.python.org/downloads/)
2. [pandas](https://pandas.pydata.org/pandas-docs/stable/install.html)
3. [python-docx](https://python-docx.readthedocs.io/en/latest/user/install.html#install)

## input files

2 input files are required

1. input excel (xlsx) file with a list of names
2. the style docx file that specifies the format of the final docx file

The only data in the input excel file is a list of names. These will be added row by row to the output docx file. For example. Name1 goes in row1 col1, name2 row1 col2, name3 row2 col1, etc. 

__Make sure that the names in the excel doc are in the sheet called Sheet1__
__Include a "Preferred_Name" column for the name you want listed in the docx file__

Also note that there cannot be spaces in file names. 

I wouldn't mess with the style docx file. The outfile consists of an invisible table with 3 columns. To change spacing between your two columns (in col1 and 3), just adjust the table.

If you want different signature line length or distance between entries, you could alter the python script (just adjust the numer of `_____`), or tell me and I can do it for you. 

