## usage:
requirement:
```
python3
pandas
re
json
```

#### usage:
command line:

`python read_excel.py excel_file header_numbers column_start column_end`

**in this command:**

header_numbers = how many rows of header in this excel

**\[** column_start column_end **)**  **(exclude right side)**

if your excel only have one row header,
please input 

`python read_excel.py excel_file 1 0 max-columns+1`
