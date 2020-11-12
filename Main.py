from higher_functions import *
from lower_functions import *

import os

'''if not(os.getcwd() == "/home/dgarci23/Report-Generator"): 
	os.chdir("/home/dgarci23/Report-Generator")'''

while True:
	try:
		Pref, CheckEmail = Download_Email()
		break
	except IMAP4.error:
		pass

if CheckEmail:

    print('archivo encontrado!')

    Pref = Pref.split('-')

    print('tipo: ', Pref[0])

    if Pref[0].lower() == 'word' : size_file = DanielaWord(Pref);

    if Pref[0].lower() == 'excel' : size_file = DanielaPPTX(Pref);

    Delete_Email()

    print('email deleted')


if CheckEmail: Send_Email(Pref[0], size_file);

if CheckEmail: print('email sent')
