from sys import exit

opt = int(input(f'Vuoi calcolare:\n1) Solo le prenotazioni future\n2) Tutte le prenotazioni\n\n'))

while opt != (1, 2):
    if opt == 1:
        import prenotazioni.main_future_reservation
        exit()
    elif opt == 2:
        import prenotazioni.main_all_reservation
        exit()
    else:
        print(f'Hai sbagliato a digitare\n')
        opt = int(input(f'Vuoi calcolare:\n1) Solo le prenotazioni future\n2) Tutte le prenotazioni\n\n'))