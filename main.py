import openpyxl as xl
import easygui as gui
from os import access, R_OK, W_OK
from os.path import exists, isfile
from shutil import copyfile
from datetime import date
from config import *


def process() -> None:
    # init year and date values
    year = date.today().year
    year = str(year)[-2] + str(year)[-1]
    month = str(date.today().month).zfill(2)
    day = str(date.today().day).zfill(2)

    # ask if to start or to end
    title_start = "Continue?"
    msg_start = "continue form?"
    while gui.ynbox(title=title_start, msg=msg_start):

        # ask for date
        title_date = "Date?"
        msg_date = "enter date:"
        fields_date = ["day", "month", "year"]
        values_date = [day, month, year]
        values_date = gui.multenterbox(title=title_date, msg=msg_date, fields=fields_date, values=values_date)
        if not values_date:
            continue
        day, month, year = values_date
        year = str(year).zfill(2)
        month = str(month).zfill(2)
        day = str(day).zfill(2)

        # get workbook file (or create it)
        wb_filename = get_workbook(base_file=monthly_filename, year=year, month=month)
        if not (
            exists(wb_filename) or
            isfile(wb_filename)
        ):
            wb_base_filename = get_workbook(base_file=base_filename)
            print("copying file {} to {}".format(wb_base_filename, wb_filename))
            copyfile(wb_base_filename, wb_filename)

        # check if workbook is readable
        if not(
            access(wb_filename, R_OK) or
            access(wb_filename, W_OK)
        ):
            title_error = "ERROR"
            msg_error = "error reading file {}".format(wb_filename)
            gui.msgbox(title=title_error, msg=msg_error)
            continue

        # load header
        wb = xl.load_workbook(wb_filename, keep_vba=True, keep_links=True, )

        # oad all sheets
        if isinstance(wb, xl.Workbook):
            for ws_name in wb.sheetnames:
                print("loading worksheet {}".format(ws_name))
                _ws = wb[ws_name]

        ws = wb.active
        header = get_header(worksheet=ws, row=header_row, column_start=header_column_start,
                            column_end=header_column_end)
        print("loaded {} header columns".format(len(header)))

        # starting entering actual data
        if len(header) > 0:
            sections = {}

            # group columns into sections by color
            for cell in header:

                # get section color, if possible
                color = "default"
                if hasattr(cell, "fill")\
                        and hasattr(cell.fill, "bgColor")\
                        and hasattr(cell.fill.bgColor, "rgb"):

                    if isinstance(cell.fill.bgColor.rgb, str):
                        color = cell.fill.bgColor.rgb

                # create section entry for new color, if needed
                if color not in sections.keys():
                    sections[color] = []
                sections[color].append(cell)

            print("found {} sections ({})".format(len(sections), str(sections.keys())))

            # ask for section
            title_section = "Select section"
            msg_section = "select the section of you data entry"
            choices_selection = []
            _n = 1
            for _choice in sections.keys():
                choices_selection.append("[{}]: {}".format(str(_n), str(_choice)))
                _n += 1
            choice = gui.buttonbox(title=title_section, msg=msg_section, choices=choices_selection)
            choice = str(choice).replace('[', '').replace(']', '').split(":")[1].strip()
            section = sections[choice]
            if not section:
                continue
            print("selected section with color {}".format(str(choice)))

            # ask for category
            title_category = "Select Column"
            msg_category = "select category where to enter value"
            choices_category = []
            # create category choices
            _n = 1
            for _category in section:
                choices_category.append("{}: {}".format(str(_n), str(_category.value)))
                _n += 1
            category_name = gui.choicebox(title=title_category, msg=msg_category, choices=choices_category)
            if not category_name:
                continue
            category_number = str(category_name).split(":")[0].strip()
            category_number = int(category_number)-1
            category = section[category_number]
            print("selected category #{}: {}".format(category_number, str(category.value)))

            # ask for value
            title_value = "Data Value"
            msg_value = "enter value to add for\n{}".format(category_name)
            value = gui.enterbox(title=title_value, msg=msg_value)
            if not value:
                continue
            value = str(value).strip().replace(',', '.')
            print("received value = {}".format(value))

            # calculate day row
            data_cell_row = date_row_start + int(day) - 1
            if data_cell_row <= date_row_end:

                # calculate cell
                data_cell_value = category.column_letter + str(data_cell_row)
                print("using data cell {}".format(str(data_cell_value)))

                # add value
                data_value = ws[data_cell_value].value
                if not data_value:
                    data_value = "="

                data_value = str(data_value)
                if "=" not in data_value:
                    data_value = "=" + data_value

                data_value += "+{}".format(str(value))
                print("new value for {} is '{}'".format(data_cell_value, data_value))
                ws[data_cell_value].value = data_value

                # save workbook
                wb.save(wb_filename)


def get_header(worksheet: xl.Workbook, row=1, column_start=2, column_end=100) -> list:
    header = []
    for column in range(column_start, column_end + 1):
        cell = worksheet.cell(row=row, column=column)
        if cell.value and len(str(cell.value).strip()) > 0:
            header.append(cell)
    return header


def get_workbook(base_file: str, year: str = "", month: str = "") -> str:
    return base_file.replace("%y", str(year)).replace("%d", str(month))


# ---- MAIN ----
if __name__ == '__main__':
    process()
