# Excel Data Pairer

Excel Data Pairer is a powerful Python package designed to efficiently extract and manage data pairs from Excel files. It's particularly useful for tasks like managing translations, key-value pairs, or any scenario where you need to extract multiple lists of items into JSON format.

## 🚀 Key Features

- **Automatic Data Pair Extraction**: Easily define and extract data pairs from Excel sheets.
- **Schema-Based Navigation**: Navigate and manage Excel files using a structured schema.
- **JSON Serialization**: Save extracted data and schemas in JSON format for easy integration with other systems.
- **Autosave Functionality**: Automatically save configurations to prevent data loss.

## 📦 Installation

Install Excel Data Pairer directly from the GitHub repository:

```bash
pip install git+https://github.com/laelhalawani/excel_data_pairer.git
```

## 🔧 Basic Usage

Here's a quick example of how to use Excel Data Pairer to extract data pairs:

```python
from excel_data_pairer import ExcelDataPairer

# Initialize the ExcelDataPairer
pairer = ExcelDataPairer(file_path='path/to/your/excel_file.xlsx')

# Add a sheet to the schema
pairer.add_sheet('Translations')

# Add a data pair (e.g., English to French translations)
pairer.add_data_pair(
    sheet_id='Translations',
    src_columns_range='A',  # English words in column A
    src_rows_range='1-10',  # First 10 rows
    mt_columns_range='B',   # French translations in column B
    mt_rows_range='1-10'    # First 10 rows
)

# Extract the data
data = pairer.get_all_data()
print(data)
```

## 🔍 Key Methods

### `add_data_pair(sheet_id, src_columns_range, src_rows_range, mt_columns_range, mt_rows_range)`

Define a pair of cell ranges to extract data from.

- `sheet_id`: Name or index of the sheet
- `src_columns_range`: Source column(s) range (e.g., 'A' or 'A-C')
- `src_rows_range`: Source row(s) range (e.g., '1-10')
- `mt_columns_range`: Target column(s) range
- `mt_rows_range`: Target row(s) range

### `get_all_data()`

Retrieve all defined data pairs across all sheets.

### `preview_range(sheet_id, columns_range, rows_range)`

Preview data from a specific range in a sheet.

### `save_to_file()` and `load_from_file(json_path)`

Save and load your data pair configurations for easy reuse.

## 💡 Use Cases

1. **Translations**: Extract source language and translations from Excel sheets.
2. **Key-Value Pairs**: Manage configuration data stored in Excel.
3. **Data Validation**: Compare expected vs. actual values in testing scenarios.
4. **Data Migration**: Extract structured data from Excel for import into databases or other systems.

## 🔒 Autosave Feature

Excel Data Pairer includes an autosave feature to prevent data loss:

```python
pairer = ExcelDataPairer(file_path='example.xlsx', autoload=True)
pairer.enable_autosave()
```

This automatically saves your configuration after each operation.

## 📘 Advanced Usage

For more advanced usage, including managing multiple sheets, updating cell values, and customizing the schema, please refer to the full documentation.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is licensed under the MIT License.