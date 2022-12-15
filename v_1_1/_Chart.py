from openpyxl.chart import PieChart, Reference, series
from openpyxl.chart.series import DataPoint

def adjust_column_width(startRow, startCol, maxRow, maxCol, ws):
    dims = {}
    for col_val in range(startCol, maxCol + 1):
        max_width = 8
        for row_val in range(startRow, maxRow + 1):
            cell = ws.cell(row=row_val, column=col_val)
            if cell.value:
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
                print(cell.value)
                # max_width = max(len(str(cell.value)), max_width)
    for col, value in dims.items():
        ws.column_dimensions[col].width = value + 4  # for rank only

def create_chart_for_department_details(year_chart_sheet, stats,CONTEST_NAME):
    for row in stats:
        year_chart_sheet.append(row)
    adjust_column_width(startRow=1, startCol=1, maxRow=len(stats), maxCol=len(stats[0]), ws=year_chart_sheet)
    pie = PieChart()
    labels = Reference(year_chart_sheet, min_col=1, min_row=2, max_row=len(stats))
    data = Reference(year_chart_sheet, min_col=4, min_row=1, max_row=len(stats))
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "ATTENDANCE PERCENTAGE " + CONTEST_NAME

    # slice = DataPoint(idx=0, explosion=20)
    # pie.series[0].data_points = [slice]

    year_chart_sheet.add_chart(pie, "F4")

    pie.height = 12.43
    pie.width = 13.48