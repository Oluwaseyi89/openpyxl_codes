from openpyxl import load_workbook
from trim_blank_rows import trim_blank_rows
from filter_func import filter_rows_by_values
from remove_duplicate_rows import remove_duplicate_rows
from filter_subset_data import filter_out_subset

filter_params = ["NEW BUSINESS", "RENEWAL"]


input_file = "./input/input_file.xlsx"
output_file = "./output/my_output.xlsx"
subset_file = "./output/subset_file.xlsx"
subset_result_file = "./output/subset_filter_result.xlsx"

key_cols = ['INSURED NAME', 'CLIENT TYPE', 'POLICY NUMBER', 'TRANS TYPE', 'CURR', 'GRP CLASS', 'SUB CLASS']


trim_blank_rows(input_file, input_file)

filter_rows_by_values(input_file, input_file, "TRANS TYPE", filter_params)

remove_duplicate_rows(input_file, output_file)

filter_out_subset(output_file, subset_file, subset_result_file, key_cols)






