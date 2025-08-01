"""Excel parser for remove unused columns."""
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, DEFAULT_FONT
from openpyxl.utils import column_index_from_string
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class ExcelParser:
    def __init__(self, workbook_path: str, rate: int, exchange_rate: float):
        """Initialize ExcelParser with the path to the workbook."""
        self.workbook_path = workbook_path
        self.rate = rate
        self.exchange_rate = exchange_rate

    def remove_unused_columns(self):
        """Remove specified unused columns from the active worksheet."""
        # load existing workbook
        wb = load_workbook(self.workbook_path)
        ws = wb.active  # Get the active worksheet

        # define column letters to be removed
        column_letters = ["A", "C", "D", "F", "H", "I", "J", "L", "M", "N", "O", "P", "Q", "S", "T", "U", "V", "W", "X", "Y", "Z"]

        # convert column letters to 1-based column indices
        column_indices = [column_index_from_string(col) for col in column_letters]

        # delete columns in reverse order to avoid index shifting
        for col_idx in sorted(column_indices, reverse=True):
            ws.delete_cols(col_idx)

        # save the modified workbook to a new file
        wb.save(self.workbook_path)

        return wb.path

    def generate_financial_report(self, output_path: str = "financial_report.xlsx"):
        """Generate structured financial report grouped by project with calculated fields."""
        wb: Workbook = load_workbook(self.workbook_path)
        ws: Worksheet = wb.active

        rows: list = list(ws.iter_rows(values_only=True))
        header, data = rows[0], rows[1:]
        data.sort(key=lambda row: row[0])

        grouped: dict = {}
        for row in data:
            project = row[1] # project row
            if project not in grouped:
                grouped[project] = []

            grouped[project].append(row)

        new_wb: Workbook = Workbook()
        new_ws = new_wb.active

        # set default font text
        DEFAULT_FONT.name = "Ubuntu"

        # Set column widths
        column_widths = [15, 45, 55, 15, 15, 15, 20, 15]
        for col_idx, width in enumerate(column_widths, start=1):
            col_letter = new_ws.cell(row=1, column=col_idx).column_letter
            new_ws.column_dimensions[col_letter].width = width

        formatted_header: list = [
            "Проект", "Ссылка на TeamWork", "Описание", "Кол-во часов",
            "Рейт/час [$]", "Курс [$]", "Стоимость (ГРН)", "Дата"
        ]

        row_cursor: int = 1

        # write header with bold text and grey background
        header_font = Font(name=DEFAULT_FONT.name, bold=True)
        header_fill = PatternFill(fill_type="solid", start_color="BDBDBD", end_color="BDBDBD")

        for col_idx, val in enumerate(formatted_header, start=1):
            cell = new_ws.cell(row=row_cursor, column=col_idx, value=val)
            cell.font = header_font
            cell.fill = header_fill

        row_cursor += 1

        for project, row_list in grouped.items():

            # merge cells from A to H (1 to 8) for project title row
            new_ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=8)
            project_cell = new_ws.cell(row=row_cursor, column=1, value=project)

            # apply bold font and background color #cfe2f3
            project_cell.font = Font(name=DEFAULT_FONT.name, bold=True)
            project_cell.fill = PatternFill(fill_type="solid", start_color="CFE2F3", end_color="CFE2F3")

            row_cursor += 1

            for row in row_list:
                date_val = row[0]
                desc = row[2] or ""
                task = row[3] or ""
                hours = float(row[4])
                task_id = row[5]

                if isinstance(date_val, str):
                    try:
                        date_val = datetime.fromisoformat(date_val)
                    except ValueError:
                        date_val = None

                total_uah: float = round(hours * self.rate * self.exchange_rate, 2)
                teamwork_link = f"https://avada.teamwork.com/#tasks/{task_id}"

                # write data to cells
                new_ws.cell(row=row_cursor, column=1, value=task_id)
                new_ws.cell(row=row_cursor, column=2, value=teamwork_link)
                new_ws.cell(row=row_cursor, column=3, value=desc)
                new_ws.cell(row=row_cursor, column=4, value=hours)
                new_ws.cell(row=row_cursor, column=5, value=self.rate)
                new_ws.cell(row=row_cursor, column=6, value=self.exchange_rate)
                new_ws.cell(row=row_cursor, column=7, value=total_uah)
                new_ws.cell(row=row_cursor, column=8, value=date_val.strftime("%d.%m.%Y"))

                row_cursor += 1

            # 3 blank lines between project groups
            row_cursor += 3

        new_wb.save(output_path)
        return output_path


