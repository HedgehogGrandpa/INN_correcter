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
    parser.add_argument('--i', metavar='INN', dest='inn', action='store', nargs='*', type=int, required=True,
                        help='������ ��������, � ������� ���� ���')
    parser.add_argument('--k', metavar='KPP', dest='kpp', action='store', nargs='*', type=int,
                        help='������ ��������, � ������� ���� ���.\n'
                             '������ �������� ��������� ��������������� ������� �������� � ���')

    args = parser.parse_args()
    return args.filename[0], args.outputFileName, args.inn, args.kpp


def main():
    filename, output_file, inns, kpps = argument_parse()
    print(filename, output_file, inns, kpps)
    formatter = VATFormatter.VATFormatter(filename, inns, kpps, output_file)
    formatter.correct_type_of_vat()


if __name__ == '__main__':
    main()
