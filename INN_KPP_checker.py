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
    parser.add_argument('-p', metavar='prefix', dest='pre', nargs='*', action='store', type=str,
                        help='������ �������� � �������� ��� ����������.\n'
                             '�������� � ������� <�����_�������1> <�������1> <�����_�������2> <�������2> ...')
    parser.add_argument('-s', metavar='suffix', dest='suf', nargs='*', action='store', type=str,
                        help='������ �������� � �������� ��� ����������.\n'
                             '�������� � ������� <�����_�������1> <�������1> <�����_�������2> <�������2> ...')

    args = parser.parse_args()
    return args.filename[0], args.outputFileName, args.inn, args.kpp, args.name, args.pre, args.suf


def correcting(ifn, inns, kpps, ofn, name, pre, suf):
    formatter = VATFormatter.VATFormatter(ifn, inns, kpps, ofn, name, pre, suf)
    formatter.correct_vat()


def gui():
    pass


def main():
    filename, output_file, inns, kpps, name, pre, suf = argument_parse()
    # correcting(filename, inns, kpps, output_file, name, pre, suf)
    formatter = VATFormatter.VATFormatter(filename, inns, kpps, output_file, name, pre, suf)
    formatter.correct_vat()


if __name__ == '__main__':
    main()
