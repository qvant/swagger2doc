# swagger2doc
simple command line tool for generation MS Word documentation for swagger API descrioption. Require python 3, python-docx, pyaml and argparse.

## Installation:
1. copy files
2. pip install pyayml
3. pip install python-docx
4. pip install argparse

## Command line arguments:  
	-i - input swagger yaml  
	-o - output docx file  
	-l - language file, that contatins words, describing API in your language in yaml format  
	-e - encoding of language file  
By default this tool use english language settings from config/en_lang.yaml and generate exapmple for swagger demo doc.