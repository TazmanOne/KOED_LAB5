import xlrd
from veryprettytable import VeryPrettyTable
import sys
import io

x = []
table = []
alfa = []
St_array = []
for i in range(1, 10):
    alfa.append(i * 0.1)
t = list(range(0, 30))
col_names = ["Неделя", "Значение"]
col_names1 = ["Альфа", "Значение"]
title = "Исходные данные"
title1 = "Выходные  данные"


def read_workbook(filename, st_col, st_end):
    read_book = xlrd.open_workbook(filename, formatting_info=True)
    sheet = read_book.sheet_by_index(1)

    for rownum in range(1, sheet.nrows):
        row = sheet.row_values(rownum, st_col, st_end)
        if "" not in row and "..." not in row:
            float_row = list(map(lambda x: float(x), row))
            table.append(float_row)
    return table


def calc_so(table):
    sum_S0 = 0
    for i in table:
        x.append(i[1])
    for j in range(1, 5):
        sum_S0 += x[j]
    return sum_S0


def print_table(table, col_names, title):
    result = VeryPrettyTable()
    result.title = title
    result.field_names = col_names
    for row in table:
        result.add_row([format(i, ".2f") for i in row])

    print(result)


def print_result_of_tab(table, col_names):
    result = VeryPrettyTable()
    tmp = []
    ll = ""
    #print(len(table))
    #print(table)
    print(result)
    for j in [0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9]:
        tmp = [elem[1] for elem in table if elem[0]==j]
        pop = []
        for elem in tmp:
            ll = "{0}".format(j)
            pop.append(elem)
            print(pop)
        result.add_column(ll, [format(i, ".2f") for i in pop])
        print(pop)
    print(result)


def my_S(alfa, t):
    s_shtrih = []
    s0 = calc_so(table)
    for alfa_s in alfa:
        beta = 1 - alfa_s
        result_s0 = (alfa_s * s0)
        s_shtrih.append([alfa_s, result_s0])
        for j in t:
            s_shtrih.append([alfa_s, alfa_s * x[j] + beta * s_shtrih[j - 1][1]])
    return s_shtrih


if __name__ == "__main__":
    print_table(read_workbook("data.xls", 0, 2), col_names, title)
    print(calc_so(table))
    print_table(my_S(alfa, t), col_names1, title1)
    print_result_of_tab(my_S(alfa, t), col_names1)
