import json
import sys
from itertools import groupby

import xlrd


def equal_left(l, index):
    for i in l[index::-1]:
        if i != '':
            return i


def construct_header(header_list: list) -> list:
    new_headers = []
    for i in range(len(header_list)):
        new_header_list = []
        for j in range(len(header_list[0])):
            if i == 0:
                new_header_list.append(equal_left(header_list[i], j))
            elif header_list[i][j] == '':
                new_header_list.append(header_list[i - 1][j]) if header_list[i - 1][j] != '' \
                    else new_header_list.append(equal_left(header_list[i], j))
            else:
                new_header_list.append(header_list[i][j])
        new_headers.append(new_header_list)
    return new_headers


def construct_json_header(headers):
    s = ''
    for i in range(len(headers)):
        if i == 0:
            s = '{'
            s += '\"' + headers[i] + '\":'
        elif headers[i] == headers[i - 1]:
            continue
        else:
            s += '{' + '\"' + headers[i] + '\":'
    return s


def construct_json_with_value(header, row_values):
    json_return = header.copy()
    for i in range(len(header)):
        json_return[i].append(row_values[i])
    return json_return


def final(json_list):
    output = []
    for i in json_list:
        parent, current = output, output
        for index, part in enumerate(i):
            for item in current:
                if part in item:
                    parent, current = current, item[part]
                    break
            else:
                if index == len(i) - 1:
                    current.append(part)
                else:
                    new_list = []
                    current.append({part: new_list})
                    parent, current = current, new_list
    return output


def read_excel(file='test.xlsx', header=1, columns=[0, 0]):
    excel = xlrd.open_workbook(file)
    sheet = excel.sheet_by_index(0)
    sheet_rows, sheet_cols = sheet.nrows, sheet.ncols
    headers = [[sheet.cell_value(j, i) for i in range(sheet_cols)] for j in range(header)]
    print(sheet_rows, sheet_cols)

    #     new_headers = construct_header(headers)

    #     headers_to_join = []
    #     for i in range(len(new_headers[0])):
    #         headers_to_join.append([new_headers[j][i] for j in range(len(new_headers))])

    #     headers_to_join_rm_duplicate = []
    #     for i in headers_to_join:
    #         headers_to_join_rm_duplicate.append(list(j for j, x in groupby(i)))
    #     print(headers_to_join,'\n',headers_to_join_rm_duplicate,'\n')

    for j in range(header, sheet_rows):
        new_headers = construct_header(headers)

        headers_to_join = []
        for i in range(len(new_headers[0])):
            headers_to_join.append([new_headers[j][i] for j in range(len(new_headers))])

        headers_to_join_rm_duplicate = []
        for i in headers_to_join:
            headers_to_join_rm_duplicate.append(list(j for j, x in groupby(i)))
        #         print(headers_to_join,'\n',headers_to_join_rm_duplicate,'\n')

        json_with_value_list = construct_json_with_value(headers_to_join_rm_duplicate, sheet.row_values(j))
        print(json_with_value_list)

        print(json.dumps(final(json_with_value_list)))

#         # 比较蠢
#         for j in range(header, 5):
#             s_json = ''
#             for i in range(len(headers_to_join)):
#                 s_json += construct_json_header(headers_to_join[i])
# #                 print(sheet.cell_value(j,i))
#                 s_json += '\"' + str(sheet.cell_value(j,i)) + '\"}'
#             print(s_json,'\n')

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("please input the file name, header range (default 0 0 means only use the 0th row as header), ")
    else:
        read_excel(sys.argv[1], int(sys.argv[2]))
