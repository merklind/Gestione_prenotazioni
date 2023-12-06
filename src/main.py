from sys import exit

from main_all_reservation import all_reservation
from main_future_reservation import future_reservation

opt = int(input(f'Vuoi calcolare:\n1) Solo le prenotazioni future\n2) Tutte le prenotazioni\n\n'))

while opt != (1, 2):
    if opt == 1:
        future_reservation()
        exit(0)
    elif opt == 2:
        all_reservation()
        exit(0)
    else:
        print(f'Hai sbagliato a digitare\n')
        opt = int(input(f'Vuoi calcolare:\n1) Solo le prenotazioni future\n2) Tutte le prenotazioni\n\n'))
