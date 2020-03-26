from datetime import date
from json import load

from prenotazioni.constants import *
from prenotazioni.utils.excel_utils import find_max_row, open_workbook, open_worksheet, duplicate_worksheet, \
    get_reservation_by_year, insert_new_dettaglio_prezzo_reservation
from prenotazioni.utils.path_utils import get_folder_path
from prenotazioni.utils.json_utils import reset_max_row_json

if __name__ == '__main__':
    folder_path = get_folder_path()
    dettaglio_prezzi_path = folder_path.joinpath(DETTAGLIO_FILE)
    wb_master_path = folder_path.joinpath(MASTER_FILE)
    wb_dettaglio_prezzi, file_path = open_workbook(dettaglio_prezzi_path)
    wb_master, useless = open_workbook(wb_master_path, True)
    ws_master = open_worksheet(wb_master, RIEPILOGO_WS)

    column_label_json = open(folder_path.joinpath(COLUMN_FILE), mode='r')
    apartment_json = open(folder_path.joinpath(APARTMENT_FILE), mode='r')
    max_row_dettaglio_prezzo_json = open(folder_path.joinpath(MAX_ROW_FILE), mode='r')
    column_label = load(column_label_json)
    apartment = load(apartment_json)
    max_row_dettaglio_prezzo = load(max_row_dettaglio_prezzo_json)

    max_row = find_max_row(ws_master)
    index_res = 1
    current_year = date.today().year


    # option = int(input(f'Vuoi ricreare il dettaglio prezzi per:\n1) Tutti gli anni\n2) L\'anno corrente\n'))
    #
    # # to create table for all reservation
    # if option == 1:
    #     reset_max_row_json()
    #     for year in range(2017, 2021):
    #         all_reservation = get_reservation_by_year(ws_master, year, max_row, column_label)
    #         ws = duplicate_worksheet(wb_dettaglio_prezzi, year)
    #         for reservation in all_reservation:
    #             print(f'Compute reservation {index_res}: {reservation["name_guest"]}')
    #             year = reservation['check-in'].year
    #             house = reservation['house'].lower()
    #
    #             column_price = apartment[house]['Dettaglio prezzi']
    #             starting_row = max_row_dettaglio_prezzo[f'{year}'][house]
    #             row_update = insert_new_dettaglio_prezzo_reservation(ws, reservation, starting_row, column_price)
    #
    #             max_row_dettaglio_prezzo[f'{year}'][house] = row_update
    #             index_res += 1
    #
    #
    # # to create table only for current year
    # elif option == 2:

    ws = duplicate_worksheet(wb_dettaglio_prezzi, current_year)
    all_reservation = get_reservation_by_year(ws_master, current_year, max_row, column_label)

    for reservation in all_reservation:
        if reservation['check-in'].year >= current_year:
            print(f'Compute reservation {index_res}: {reservation["name_guest"]}')
            house = reservation['house'].lower()


            column_price = apartment[house]['Dettaglio prezzi']
            starting_row = max_row_dettaglio_prezzo[f'{current_year}'][house]
            row_update = insert_new_dettaglio_prezzo_reservation(ws, reservation, starting_row, column_price)

            max_row_dettaglio_prezzo[f'{current_year}'][house] = row_update

            index_res += 1


    wb_dettaglio_prezzi.save(file_path)
    wb_dettaglio_prezzi.close()

    # save_max_row_json(max_row_dettaglio_prezzo)
