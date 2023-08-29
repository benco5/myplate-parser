def ordinal_suffix(i):
    """
    Returns appropriate ordinal suffix for integer input (-st, -nd, -rd, -th)

    Parameters
    ----------
    i : integer

    Returns
    -------
    string
    """
    return "th" if i in {11, 12, 13} else {1: "st", 2: "nd", 3: "rd"}.get(i % 10, "th")


def formatted_datestring(dt):
    """
    Returns formatted datestring like 'January 2nd, 2023' to match the date values
    format that appears in the MyPlate .xls export.

    Parameters
    ----------
    dt : datetime object

    Returns
    -------
    string
    """
    return f"{dt:%B} {dt.day}{ordinal_suffix(dt.day)}, {dt.year}"


def write_sheet_data(sheet, data):
    """
    Writes rows to worksheet in xlwt (.xls) workbook

    Parameters
    ----------
    sheet : xlwt sheet object
    data  : 2d list of cell value data

    Returns
    -------
    None
    """
    for row in range(len(data)):
        for column, value in enumerate(data[row]):
            sheet.write(row, column, value)
