#!/usr/bin/env python
# -*- coding: cp1251 -*-

import argparse
import xlrd
import xlsxwriter


def argument_parse():
    parser = argparse.ArgumentParser(
        description='��������� ��������� ���������� ���-���, ���� ��������, �� ������ '
                    '(���� ������ ����������� ������� �����)')
    parser.add_argument('filename', nargs=1,
                        help='��� ����� .xls ��� .xlsx ������� � ����������� �� ������������')
    parser.add_argument('--INN', dest='INNs', action='store', nargs='*', type=int, required=True,
                        help='������ ��������, � ������� ���� ���')
    parser.add_argument('--KPP', dest='KPPs', action='store', nargs='*', type=int,
                        help='������ ��������, � ������� ���� ���.\n'
                             '������ �������� ��������� ��������������� ������� �������� � ���')
    args = parser.parse_args()
    if args.KPPs is None:
        args.KPPs = [None] * len(args.INNs)
    else:
        args.KPPs.extend([None] * (len(args.INNs) - len(args.KPPs)))
    inn_kpp = {inn: kpp for inn, kpp in zip(args.INNs, args.KPPs)}
    return args.filename[0], inn_kpp


def read_sheet(sheet, inn_kpp):
    for row in sheet.get_rows():
        for inn in inn_kpp:
            # ���� ����� ��� ������ ���
            # ������� ������ �� ���
            kpp = inn_kpp[inn]
            if not (kpp is None or row[inn].ctype == 1):
                # ���� ��� ������, �� ��� �����������
                if row[kpp].value:
                    new_kpp = '{:0>9}'.format(int(row[kpp].value))
                    new_inn = '{:0>10}'.format(int(row[inn].value))
                    row[kpp].ctype = 1
                    row[inn].ctype = 1
                    row[kpp].value = new_kpp
                    row[inn].value = new_inn
                # ����� ��� ��
                else:
                    new_inn = '{:0>12}'.format(int(row[inn].value))
                    row[inn].ctype = 1  # str
                    row[kpp].ctype = 1
                    row[inn].value = new_inn
            # ���� ��� ������ � ��� � ��������, �� ����� �������� �������� ����
            else:
                if row[inn].ctype == 2:  # ���� �����
                    l = len(str(int(row[inn].value)))
                    new_inn = int(row[inn].value)
                    if l == 9:
                        new_inn = '{:0>10}'.format(new_inn)
                    if l == 11:
                        new_inn = '{:0>12}'.format(new_inn)
                    row[inn].ctype = 1
                    row[inn].value = new_inn
        new_row = [x.value for x in row]
        yield new_row


def main():
    filename, inn_kpp = argument_parse()
    with xlrd.open_workbook(filename) as read_book:
        sheet = read_book.sheet_by_index(0)
        with xlsxwriter.Workbook('Corrected_' + filename) as write_book:
            write_sheet = write_book.add_worksheet('Corrected')
            row_num = 0
            for row in read_sheet(sheet, inn_kpp):
                col_num = 0
                for cell in row:
                    write_sheet.write_string(row_num, col_num, cell)
                    col_num += 1
                row_num += 1

if __name__ == '__main__':
    main()
