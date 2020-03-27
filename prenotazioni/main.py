from datetime import datetime
from json import load
from pathlib import PurePath

from prenotazioni.constants import MASTER_FILE, COLUMN_FILE, APARTMENT_FILE, RIEPILOGO_WS, RENDICONTO_WS
from prenotazioni.utils.excel_utils import find_max_row, get_info_reservation, open_workbook, open_worksheet
from prenotazioni.style_worksheet.style_excel import erase_future_reservation, set_rendiconto_occupation_cell,\
    set_rendiconto_gross_cell, set_rendiconto_net_cell, set_cell_reservation_break
from prenotazioni.utils.json_utils import find_column_apartment
from prenotazioni.utils.path_utils import get_folder_path

# set path of the folder
folder_path = get_folder_path()


name_file = MASTER_FILE
wb_master_path = folder_path.joinpath(name_file)

# open all needed files
json_column_file = open(str(PurePath.joinpath(folder_path, COLUMN_FILE)))
json_apartment_file = open(str(PurePath.joinpath(folder_path, APARTMENT_FILE)))
apartment_json = load(json_apartment_file)
column_label = load(json_column_file)

wb, file_path = open_workbook(wb_master_path, data_only=False)
wb_only_data, useless = open_workbook(wb_master_path, data_only=True)
ws_riepilogo = open_worksheet(wb, RIEPILOGO_WS)
ws_rendiconto = open_worksheet(wb, RENDICONTO_WS)
ws_riepilogo_data_only = open_worksheet(wb_only_data, RIEPILOGO_WS)

# create a copy of the file (security purpose)
# create_copy(folder_path, name_file)

print('File caricato\n')
start_row_riepilogo = 2
end_row_riepilogo = find_max_row(ws_riepilogo)
start_row_rendiconto = 4
end_row_rendiconto = 1303
info_reservation = {}
today = datetime(datetime.today().year, datetime.today().month, datetime.today().day)


# erase_all_reservation_rendiconto(ws_rendiconto, start_row_rendiconto, end_row_rendiconto)
erase_future_reservation(ws_rendiconto, today, start_row_rendiconto, end_row_rendiconto)

# find the first row in riepilogo worksheet associated with the firts reservation of the year
while ws_riepilogo_data_only.cell(start_row_riepilogo, column_label['Entrata']).value.year != today.year:
    start_row_riepilogo += 1

#RENDICONTO WORKSHEET #
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

        if date_check_in >= today:

            # set thick border for range cell, set value to 1 and set background color
            set_rendiconto_occupation_cell(ws_rendiconto, start_row_rendiconto, info_reservation, taken_column)
            # set the price per night in Rendiconto worksheet
            set_rendiconto_gross_cell(ws_rendiconto, ws_riepilogo, gross_column, start_column_price, start_row_rendiconto, number_reservation, info_reservation)
            # set the net price per night in Rendiconto worksheet
            set_rendiconto_net_cell(ws_rendiconto, start_row_rendiconto, net_column, gross_column, info_reservation)

        if date_check_in < today < date_check_out:

            set_cell_reservation_break(ws_rendiconto, start_row_rendiconto, info_reservation, taken_column, gross_column,
                                       net_column, today, info_reservation['Tipo tariffa'])

    info_reservation.clear()

wb.save(file_path)
wb.close()
