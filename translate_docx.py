# Import re for regex functions
# Import sys for getting the command line arguments
# import argparse to handle tac arguments
import argparse
import re
import sys

# Import docx to work with .docx files.
# Must be installed: pip install python-docx
from docx import Document

# Import the name mappings
from US_UK_dict import uk_us

argParser = argparse.ArgumentParser()
argParser.add_argument("-i", "--inputfile", help="Enter the input file path.")
argParser.add_argument("-l", "--language", type=str, help="Enter the language you want to change to, US or UK.")
args = argParser.parse_args()

# Check if Command Line Arguments are passed.
if len(sys.argv) < 2:
    print('Not Enough arguments where supplied')
    sys.exit()
elif not (args.language == 'UK' or args.language == 'US'):
    print('Select a supported language')
    sys.exit()

# Store file path from CL Arguments.
file_path = args.inputfile

# Create a variable based on language selected
# Defaults to translate things to UK
languageFrom = 1
languageTo = 0
if args.language == 'US':
    languageFrom = 0
    languageTo = 1
elif args.language == 'UK':
    languageFrom = 1
    languageTo = 0

if file_path.endswith('.docx'):
    doc = Document(file_path)
    # Loop through replacer arguments
    occurences = {}
    for replaceArgs in uk_us[0:]:
        # split the word=replacedword into a list
        replaceArg = replaceArgs
        # initialize the number of occurences of this word to 0
        occurences[replaceArg[languageFrom]] = 0
        # Loop through paragraphs
        for para in doc.paragraphs:
            # Loop through runs (style spans)
            for run in para.runs:
                # if there is text on this run, replace it
                if run.text:
                    # get the replacement text
                    replaced_text = re.sub(replaceArg[languageFrom], replaceArg[languageTo], run.text, 999)
                    if replaced_text != run.text:
                        # if the replaced text is not the same as the original
                        # replace the text and increment the number of occurences
                        run.text = replaced_text
                        occurences[replaceArg[languageFrom]] += 1

    # print the number of occurences of each word
    for word, count in occurences.items():
        if count > 0:
            print(f"The word {word} was found and replaced {count} times.")

    # make a new file name by adding "_new" to the original file name
    new_file_path = file_path.replace(".docx", "_new.docx")
    # save the new docx file
    doc.save(new_file_path)
else:
    print('The file type is invalid, only .docx are supported')
