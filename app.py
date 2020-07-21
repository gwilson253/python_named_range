from openpyxl import load_workbook
from pandas import DataFrame

wb = load_workbook(filename='example.xlsx')


def get_named_range_cells(wb, defined_name):
    named_range = wb.defined_names[defined_name]
    sheet_name, range_str = named_range.attr_text.split('!')
    return wb[sheet_name][range_str]


def tabular_cells_to_df(cells):
    cells_col = map(list, zip(*cells))
    cells_col = [[_.value for _ in col] for col in cells_col]
    data_dict = {_[0]: _[1:] for _ in cells_col}
    return DataFrame(data_dict)


def non_tabular_cells_to_df(cells, labels):
    data_dict = {l: [c.value] for l, c in zip(labels, cells)}
    return DataFrame(data_dict)


# non-tabular named range to pandas.DataFrame
t1_names = ['table1_field1', 'table1_field2', 'table1_field3']
t1 = [get_named_range_cells(wb, n) for n in t1_names]
t1_labels = ['T1F1', 'T1F2', 'T1F3']
print('Non-tabular Named-Range Example:')
print(non_tabular_cells_to_df(t1, t1_labels))
print()

# tabular named range to pandas.DataFrame
t2 = get_named_range_cells(wb, 'table2')
print('Tabular Named-Range Example:')
print(tabular_cells_to_df(t2))
