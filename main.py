import functions as f
import easygui as gui
from datetime import date
import config as cfg


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
        values_date = f.ask_for_date(year, day, month)
        if not values_date:
            continue
        year, day, month = values_date

        # get workbook file (or create it)
        wb_data = f.get_workbook(year, month)

        # check if workbook is readable
        if not wb_data:
            title_error = "ERROR"
            msg_error = f"error reading file {month}/{year}"
            gui.msgbox(title=title_error, msg=msg_error)
            continue
        workbook_filename, workbook, active_worksheet = wb_data

        # read header
        line_data = f.get_workbook_line_data(worksheet=active_worksheet, row=cfg.header_row, column_start=cfg.header_column_start,
                                                column_end=cfg.header_column_end)
        print(f"loaded {len(line_data)} header columns")

        # starting entering actual data
        if len(line_data) > 0:
            sections = f.group_sections_by_bg_color(line_data)

            # ask for section
            section_selection = f.ask_for_section(sections)
            if not section_selection:
                continue
            choice, section = section_selection
            print(f"selected section with color {str(choice)}")

            # ask for category
            category_data = f.ask_for_category(section)
            if not category_data:
                continue
            category_number, category_name, category = category_data

            # ask for value
            value = f.ask_for_value()
            if not value:
                continue

            # calculate day row
            if f.add_value_to_worksheet(worksheet=active_worksheet,
                                      day=day,
                                      category=category,
                                      value=value):
                workbook.save(workbook_filename)
                print(f"data saved to {workbook_filename}")
            else:
                print(f"data NOT saved to {workbook_filename}")


# ---- MAIN ----
if __name__ == '__main__':
    process()
