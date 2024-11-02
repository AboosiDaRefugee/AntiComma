Excel Column Expander (AntiComma)
This Python application reads an Excel file and detects columns with cells containing values separated by a specified delimiter (e.g., a comma). For each unique separated value in these columns, the program creates a new column, indicating the presence (1) or absence (0) of that value in each row.

Features
Automatic Column Expansion: Only columns containing values separated by the specified separator will be expanded.
Customizable Separator: The user can specify the delimiter (default is a comma).

Requirements
pip install pandas openpyxl
