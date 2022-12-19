# WordFindReplace
I had an issue with Word docx files and being able to change the language all at once. 

Taking the word list from here: http://www.tysto.com/uk-us-spelling-list.html
and the base of the code: https://www.thepythoncode.com/article/replace-text-in-docx-files-using-python

I improved on this by giving the ability to change from US to UK or vice versa. The word list will do all lower or capitalised first letter replacements. It will miss words that are all upper. 

TODO: Get this working for PPTX files. 

```
usage: translate_docx.py [-h] [-i INPUTFILE] [-l LANGUAGE]

options:
  -h, --help            show this help message and exit
  -i INPUTFILE, --inputfile INPUTFILE
                        Enter the input file path.
  -l LANGUAGE, --language LANGUAGE
                        Enter the language you want to change to, US or UK.
```
