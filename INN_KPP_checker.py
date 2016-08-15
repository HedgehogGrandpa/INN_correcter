# -*- coding: cp1251 -*-
import argparse
import VATFormatter


def argument_parse():
    parser = argparse.ArgumentParser(
        description='��������� ��������� ���������� ���-���, ���� ��������, �� ������ '
                    '(���� ������ ����������� ������� �����)')
    parser.add_argument('filename', nargs=1,
                        help='��� ����� .xls ��� .xlsx ������� � ����������� �� ������������')
    parser.add_argument('outputFileName', nargs='?', default='',
                        help='��� ��������� �����')
    parser.add_argument('-i', metavar='INN', dest='inn', action='store', nargs='+', type=int, required=True,
                        help='������ ��������, � ������� ���� ���')
    parser.add_argument('-k', metavar='KPP', dest='kpp', action='store', nargs='*', type=int,
                        help='������ ��������, � ������� ���� ���.\n'
                             '������ �������� ��������� ��������������� ������� �������� � ���')
    parser.add_argument('-n', metavar='name_column', dest='name', nargs='*', action='store', type=int,
                        help='������ ��������, ������������ ��� ������������ ����� �������� ������ ��� ��\n'
                             '������� ����� ����������� � ��������� ������')

    args = parser.parse_args()
    return args.filename[0], args.outputFileName, args.inn, args.kpp, args.name


def correcting(ifn, inns, kpps, ofn, name):
    formatter = VATFormatter.VATFormatter(ifn, inns, kpps, ofn, name)
    formatter.correct_type_of_vat()



def gui():
    pass


def main():
    filename, output_file, inns, kpps, name = argument_parse()
    correcting(filename, inns, kpps, output_file, name)


if __name__ == '__main__':
    main()
