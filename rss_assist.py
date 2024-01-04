import time
import xlwings
import xlwings as xw

STOCK_CODE_COLUMN = "A"
STOCK_NAME_COLUMN = "B"

TARGET_COLUMN_NAMES = ["現在値", "売買圧力比率"]

UPDATE_INTERVAL_SECONDS = 10

UP_COLORS = ["#edf5fb", "#c4ddf4", "#9bc6ec", "#71afe5", "#4897dd", "#2580ce"]
DOWN_COLORS = ["#fbedee", "#f3c4c5", "#ec9b9d", "#e47275", "#dc494d", "#cd272b"]


class Record:
    def __init__(self, value, timestamp):
        assert type(value) != type(None)
        self.value = value
        self.timestamp = timestamp

    def __repr__(self):
        return str(self.value)


class Property:
    def __init__(self, name: str, column: str, record: Record):
        self.name = name
        self.column = column
        self.records = [record]
        self.updated = False

    def __repr__(self):
        return "{}, {}".format(self.name, self.records)

    def get_column_letter(self):
        return self.column


class Stock:
    def __init__(self, name: str, code: int, row: int):
        self.name: str = name
        self.code = code
        self.row = row
        self.properties = []
        self.updated = False

    def __repr__(self):
        return "({}, {}, {})".format(self.get_row_letter(), self.name, self.properties)

    def get_row_letter(self):
        return str(1 + self.row)


def get_column_letter(column: int):  # zero origin
    column_letter = ""
    while column > 25:
        column_letter += chr(ord("A") + int(column / 26) - 1)
        column = column - (int(column / 26)) * 26
    column_letter += chr(ord("A") + (int(column)))
    return column_letter


def find_column_by_name(name: str, row="1", maxColumns=256):
    sheet = xw.books.active.sheets.active
    if sheet is None:
        return None
    for i in range(0, maxColumns):
        column = get_column_letter(i)
        if sheet[column + row].value == name:
            return column
    return None


def get_last_of(elements):
    assert len(elements) >= 1
    return elements[len(elements) - 1]


def get_before_last_of(elements):
    assert len(elements) >= 2
    return elements[len(elements) - 2]


def get_table_size(sheet: xlwings.main.Sheet):
    last_cell = sheet["A1"].current_region.last_cell
    return (last_cell.column, last_cell.row)


def update(sheet, stocks: list, table_size: tuple):
    timestamp = time.time()

    row_from = "1"
    row_to = str(table_size[1])

    column_letters = []
    for target_column_name in TARGET_COLUMN_NAMES:
        column_letters.append(find_column_by_name(target_column_name))

    records = []
    for i, target_column_name in enumerate(TARGET_COLUMN_NAMES):
        column_letter = column_letters[i]
        record = sheet[column_letter + row_from + ":" + column_letter + row_to].value
        records.append(record)

    for stock in stocks:
        stock.updated = False
        for i, record in enumerate(records):
            if stock.row >= len(record):
                continue
            value = record[stock.row]
            if value is None:
                print("[ERROR] value is None -", stock.name, stock.get_row_letter())
            if len(stock.properties) <= i:
                stock.properties.append(
                    Property(
                        TARGET_COLUMN_NAMES[i],
                        column_letters[i],
                        Record(value, timestamp),
                    )
                )
                stock.properties[i].updated = True
            elif get_last_of(stock.properties[i].records).value != value:
                stock.properties[i].records.append(Record(value, timestamp))
                stock.properties[i].updated = True
            else:
                stock.properties[i].updated = False


def clear_background_colors(sheet, table_size: tuple):
    for target_column_name in TARGET_COLUMN_NAMES:
        colomn_letter = find_column_by_name(target_column_name)
        for row in range(table_size[1]):
            sheet[colomn_letter + str(1 + row)].color = "#ffffff"


def main():
    sheet = xw.books.active.sheets.active
    if sheet is None:
        return
    table_size = get_table_size(sheet)

    row_from = "1"
    row_to = str(table_size[1])

    stocks: list = []
    names = sheet[STOCK_NAME_COLUMN + row_from + ":" + STOCK_NAME_COLUMN + row_to].value
    codes = sheet[STOCK_CODE_COLUMN + row_from + ":" + STOCK_CODE_COLUMN + row_to].value
    for i in range(0, min(len(names), len(codes))):
        name = names[i]
        code = codes[i]
        if name is None or code is None:
            continue
        code = int(code)
        stocks.append(Stock(name, code, i))

    clear_background_colors(sheet, table_size)

    last_update_time_seconds = 0
    # main loop
    while True:
        current_time_seconds = time.time()
        if current_time_seconds - last_update_time_seconds >= UPDATE_INTERVAL_SECONDS:
            update(sheet, stocks, table_size)
            for stock in stocks:
                for property in stock.properties:
                    if property.updated == False or len(property.records) <= 1:
                        continue
                    # print("[DEBUG]", "updated -", property)

                    records = property.records
                    if (
                        get_before_last_of(records).value < get_last_of(records).value
                    ):  # increase
                        value = get_last_of(records).value
                        count = 0
                        for i in reversed(range(len(records) - 1)):
                            if records[i].value > value:
                                break
                            elif records[i].value < value:
                                count += 1
                            value = records[i].value
                        color = UP_COLORS[min(count, len(UP_COLORS) - 1)]
                        sheet[
                            property.get_column_letter() + stock.get_row_letter()
                        ].color = color
                        # print("[DEBUG]", "UP" + stock.name + " - ", count)
                    elif (
                        get_before_last_of(records).value > get_last_of(records).value
                    ):  # decrease
                        value = get_last_of(records).value
                        count = 0
                        for i in reversed(range(len(records) - 1)):
                            if records[i].value < value:
                                break
                            elif records[i].value > value:
                                count += 1
                            value = records[i].value
                        color = DOWN_COLORS[min(count, len(DOWN_COLORS) - 1)]
                        sheet[
                            property.get_column_letter() + stock.get_row_letter()
                        ].color = color
                        # print("[DEBUG]", "DOWN" + stock.name + " - ", count)
            last_update_time_seconds = current_time_seconds

        time.sleep(0.5)


if __name__ == "__main__":
    main()
