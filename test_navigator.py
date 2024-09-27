from excel_navigator import ExcelNavigator


nav  = ExcelNavigator(autoload=True)
print(nav.list_excel_files('./excel_files'))
nav.select_excel_file(dir_path='./excel_files', file_name='example.xlsx')
print(nav.list_file_sheets())
nav.add_sheet('DFW_new_categories')
print(nav.preview_range('DFW_new_categories', 'A-C', '1-1'))