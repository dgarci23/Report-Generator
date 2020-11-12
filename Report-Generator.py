from higher_functions import *
from lower_functions import *

while True:

    Pref, CheckEmail = Download_Email()

    if CheckEmail:

        print('archivo encontrado!')

        Pref = Pref.split('-')

        print('tipo: ', Pref[0])

        if Pref[0].lower() == 'word' : size_file = DanielaWord(Pref);

        if Pref[0].lower() == 'excel' : size_file = DanielaPPTX(Pref);



    if CheckEmail: Send_Email(Pref[0], size_file);

    if CheckEmail: Delete_Email()

    if CheckEmail: print('email sent')

    time.sleep(120)
