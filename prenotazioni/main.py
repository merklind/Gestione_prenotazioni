opt = int(input(f'Vuoi calcolare:\n1) Solo le prenotazioni future\n2) Tutte le prenotazioni\n'))

if opt == 1:
    import prenotazioni.main_future_reservation
elif opt == 2:
    import prenotazioni.main_all_reservation
else:
    print(f'Hai sbagliato a digitare')