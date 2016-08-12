# -*- coding: cp1251 -*-
import argparse
import VATFormatter
import sys
from PyQt5.QtWidgets import QApplication, QWidget


def argument_parse():
    parser = argparse.ArgumentParser(
        description='��������� ��������� ���������� ���-���, ���� ��������, �� ������ '
                    '(���� ������ ����������� ������� �����)')
    parser.add_argument('filename', nargs=1,
                        help='��� ����� .xls ��� .xlsx ������� � ����������� �� ������������')
    parser.add_argument('outputFileName', nargs='?', default='',
                        help='��� ��������� �����')
    parser.add_argument('-i', metavar='INN', dest='inn', action='store', nargs='*', type=int, required=True,
                        help='������ ��������, � ������� ���� ���')
    parser.add_argument('-k', metavar='KPP', dest='kpp', action='store', nargs='*', type=int,
                        help='������ ��������, � ������� ���� ���.\n'
                             '������ �������� ��������� ��������������� ������� �������� � ���')

    args = parser.parse_known_args()
    args, er = args[0], args[1]
    return (args.filename[0], args.outputFileName, args.inn, args.kpp), er


def correcting(ifn, inns, kpps, ofn):
    try:
        formatter = VATFormatter.VATFormatter(ifn, inns, kpps, ofn)
        formatter.correct_type_of_vat()
    except Exception as e:
        print(e)

def gui():
    pass


def main():
    (filename, output_file, inns, kpps), er = argument_parse()
    gui() if er else correcting(filename, inns, kpps, output_file)

    #formatter = VATFormatter.VATFormatter(filename, inns, kpps, output_file)
    #formatter.correct_type_of_vat()


if __name__ == '__main__':
    main()
