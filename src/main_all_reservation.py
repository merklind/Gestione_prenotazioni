from datetime import datetime
from json import load

from src.constants import MASTER_FILE, COLUMN_FILE, APARTMENT_FILE, RIEPILOGO_WS, RENDICONTO_WS
from src.style_worksheet.style_excel import set_rendiconto_occupation_cell, \
    set_rendiconto_gross_cell, set_rendiconto_net_cell, erase_all_reservation_rendiconto
from src.utils.excel_utils import find_max_row, get_info_reservation, open_workbook, open_worksheet
from src.utils.json_utils import find_column_apartment
from src.utils.path_utils import create_copy, build_file_path


def all_reservation():
    wb_master_path = build_file_path(MASTER_FILE)
    create_copy(wb_master_path)

    column_file = open(build_file_path(COLUMN_FILE))
    apartment_file = open(build_file_path(APARTMENT_FILE))

    # open all needed files
    column_label = load(column_file)
    apartment_json = load(apartment_file)

    print(f'Apertura file excel in corso...\n')

    wb, file_path = open_workbook(wb_master_path, data_only=False)
    wb_only_data, useless = open_workbook(wb_master_path, data_only=True)

    ws_riepilogo = open_worksheet(wb, RIEPILOGO_WS)
    ws_rendiconto = open_worksheet(wb, RENDICONTO_WS)
    ws_riepilogo_data_only = open_worksheet(wb_only_data, RIEPILOGO_WS)
    ws_rendiconto_data_only = open_worksheet(wb_only_data, RENDICONTO_WS)

    print('File aperti\n')

    start_row_riepilogo = 2
    end_row_riepilogo = find_max_row(ws_riepilogo)
    start_row_rendiconto = 4
    end_row_rendiconto = ws_rendiconto_data_only.max_row
    info_reservation = {}
    today = datetime(datetime.today().year, datetime.today().month, datetime.today().day)

    erase_all_reservation_rendiconto(ws_rendiconto, start_row_rendiconto, end_row_rendiconto)

    # RENDICONTO WORKSHEET #
    # iter over all reservations in Riepilogo worksheet
    for number_reservation in range(start_row_riepilogo, end_row_riepilogo):
        print(f'Compute reservation: {number_reservation}')

        # get the state of the reservation
        state = ws_riepilogo_data_only.cell(number_reservation, column_label['Stato prenotazione']).value
        date_check_in = ws_riepilogo_data_only.cell(row=number_reservation, column=column_label['Entrata']).value
        date_check_out = ws_riepilogo_data_only.cell(row=number_reservation, column=column_label['Uscita']).value

        # find the row of the check-in reservation in Rendiconto worksheet
        while ws_rendiconto.cell(start_row_rendiconto, 1).value != date_check_in:
            start_row_rendiconto += 1

            if start_row_rendiconto == 10000:
                start_row_rendiconto = 2

        # check if the reservation isn't "Cancellata"
        if state.lower() != 'cancellata':
            info_reservation = get_info_reservation(ws_riepilogo_data_only, number_reservation, column_label)
            # find in apartment.json file the right column for the specified house and apartment for "Occupazione" subtable
            taken_column = find_column_apartment(info_reservation, apartment_json, 'Occupazione')
            # find in apartment.json file the right column for the specified house and apartment for "Incasso lordo" table
            gross_column = find_column_apartment(info_reservation, apartment_json, 'Incasso lordo')
            # find in apartment.json file the right column for the specified house and apartment for "Incasso netto" table
            net_column = find_column_apartment(info_reservation, apartment_json, 'Incasso netto')
            # find in column_label.json file the right column for the price per night
            start_column_price = column_label['Giorno 1']

            # set thick border for range cell, set value to 1 and set background color
            set_rendiconto_occupation_cell(ws_rendiconto, start_row_rendiconto, info_reservation, taken_column)
            # set the price per night in Rendiconto worksheet
            set_rendiconto_gross_cell(ws_rendiconto, ws_riepilogo, gross_column, start_column_price,
                                      start_row_rendiconto,
                                      number_reservation, info_reservation)
            # set the net price per night in Rendiconto worksheet
            set_rendiconto_net_cell(ws_rendiconto, start_row_rendiconto, net_column, gross_column, info_reservation)

        info_reservation.clear()

    print(f'\nSalvataggio file excel')

    wb.save(file_path)
    wb.close()
