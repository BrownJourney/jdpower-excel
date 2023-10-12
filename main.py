from openpyxl.utils.exceptions import InvalidFileException
from openpyxl import load_workbook
import openpyxl.utils.cell
import requests
from copy import copy

file_path = input("Enter file path:\n")


def copy_style(cell_style, new_cell):
    new_cell.font = copy(cell_style.font)
    new_cell.border = copy(cell_style.border)
    new_cell.fill = copy(cell_style.fill)
    new_cell.number_format = copy(cell_style.number_format)
    new_cell.protection = copy(cell_style.protection)
    new_cell.alignment = copy(cell_style.alignment)


def main():
    wb = None
    try:
        wb = load_workbook(file_path)
    except FileNotFoundError:
        print("This file does not exists! Check defined path again!")
        return
    except InvalidFileException:
        print("This file extension is not supported!")
        return

    ws = wb.active
    jd_clean_column = ws.max_column + 1
    jd_average_column = ws.max_column + 2
    jd_rough_column = ws.max_column + 3
    ws.insert_cols(jd_clean_column)
    ws.insert_cols(jd_average_column)
    ws.insert_cols(jd_rough_column)

    i = 0
    reserve_headers = 0

    vin_column = 'B'
    mileage_column = 'I'

    for row in ws.iter_rows():
        reserve_headers = reserve_headers + 1
        empty = True
        for cell in row:
            if cell.value is not None:
                val = cell.value.lower()
                found_vin = val.find("vin") != -1
                found_mileage = (val.find("odometer") != -1 or val.find("mileage") != -1) and val.find("unit") == -1
                if found_vin or found_mileage:
                    print(val)
                    empty = False

                if found_vin:
                    vin_column = openpyxl.utils.cell.get_column_letter(cell.column)

                if found_mileage:
                    mileage_column = openpyxl.utils.cell.get_column_letter(cell.column)

        if not empty:
            break

    jd_headers = {
        openpyxl.utils.cell.get_column_letter(jd_clean_column): {
            "name": "JD Trade Clean",
            "id": "adjustedcleantrade"
        },
        openpyxl.utils.cell.get_column_letter(jd_average_column): {
            "name": "JD Trade Average",
            "id": "adjustedaveragetrade"
        },
        openpyxl.utils.cell.get_column_letter(jd_rough_column): {
            "name": "JD Trade Rough",
            "id": "adjustedroughtrade"
        }
    }

    header_offset_y = str(reserve_headers)
    cell_style_header = ws[openpyxl.utils.cell.get_column_letter(ws.max_column) + header_offset_y]
    cell_style = ws[openpyxl.utils.cell.get_column_letter(ws.max_column) + str(reserve_headers + 1)]
    for key in jd_headers:
        ws[key + header_offset_y] = jd_headers[key]["name"]
        new_cell = ws[key + header_offset_y]
        copy_style(cell_style_header, new_cell)
        ws.column_dimensions[key].width = 20

    for cell in ws[vin_column]:
        vin = cell.value
        mileage = ws[mileage_column][i].value

        i = i + 1

        if i <= reserve_headers:
            continue

        url = ("https://cloud.jdpower.ai/data-api/UAT/valuationservices/valuation/defaultVehicleAndValuesByVin?"
               "period=0&vin={0}&region=<YOUR_CURRENT_REGION>&mileage={1}")
        response = requests.get(url.format(vin, mileage), headers={
            "api-key": "<YOUR_AUTH_TOKEN>", # Authorization token
            "accept": "application/json"
        })
        response = response.json()
        if not "result" in response:
            print("No vehicle found!")
            continue

        vehicle = response["result"][0]
        if not vehicle:
            print("Unable to find vehicle by this VIN (" + vin + ")!")
            break

        for key in jd_headers:
            ws[key + str(i)] = "$" + str(vehicle[jd_headers[key]["id"]])
            copy_style(cell_style, ws[key + str(i)])

        print("Loaded book values for VIN={0} and MILEAGE={1}".format(vin, mileage))

    file_path.replace(".xlsx", "")
    wb.save(file_path + "_bookvalues.xlsx")
    print("Table is saved!")


main()
