def is_contains_leetcode_ids(string):
    string = str(string).lower()

    # return true if cell has string leetcode id
    return string.__contains__("user") and string.__contains__("id")

def is_contains_leetcode_names(column, row,sheet_obj):
    cell = str(sheet_obj.cell(row=row, column=column).value).lower()
    return cell.__contains__("name")


def get_row_col_position_for_leetcode_id(sheet_obj):
    for row_ind in range(1, sheet_obj.max_row + 1):
        for col_ind in range(1, sheet_obj.max_column + 1):
            curr_val = sheet_obj.cell(row=row_ind, column=col_ind).value
            if curr_val is not None and is_contains_leetcode_ids(curr_val):
                return [row_ind, col_ind]
    return None

def get_row_col_position_for_leetcode_names(row_val,sheet_obj):
    for col in range(1, sheet_obj.max_column + 1):
        if is_contains_leetcode_names(col, row_val,sheet_obj):
            return col



def searchMaxCol(leet_code_id_row, startCol,sheet_obj):
    #sheet_obj represents current sheet_obj
    for col_val in range(startCol, sheet_obj.max_column + 1):
        curr_val = sheet_obj.cell(row=leet_code_id_row, column=col_val).value
        if curr_val is None:
            return col_val - 1
    return sheet_obj.max_column

def searchMaxRow(startRow, leet_code_id_col,sheet_obj):
    for row_val in range(startRow, sheet_obj.max_row + 1):
        curr_val = sheet_obj.cell(row=row_val, column=leet_code_id_col).value
        if curr_val is None:
            return row_val - 1
    return sheet_obj.max_row