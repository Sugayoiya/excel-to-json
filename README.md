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

`python read_excel.py excel_file header_start header_end column_start column_end`

in this:

[header_start, header_end]
[column_start column_end) (exclude right side)

if your excel only have one row header,
please input `python read_excel.py excel_file 0 0 0 max-columns+1`
