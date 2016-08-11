# -*- coding: cp1251 -*-
import argparse
import VATFormatter


def argument_parse():
    parser = argparse.ArgumentParser(
        description='Программа проверяет валидность ИНН-КПП, если возможно, то правит '
                    '(пока только исправление ведущих нулей)')
    parser.add_argument('filename', nargs=1,
                        help='Имя файла .xls или .xlsx формата с информацией об организациях')
    parser.add_argument('outputFileName', nargs='?', default='',
                        help='имя выходного файла')
    parser.add_argument('--i', metavar='INN', dest='inn', action='store', nargs='*', type=int, required=True,
                        help='Номера столбцов, в которых есть ИНН')
    parser.add_argument('--k', metavar='KPP', dest='kpp', action='store', nargs='*', type=int,
                        help='Номера столбцов, в которых есть КПП.\n'
                             'Номера столбцов указывать соответствующие номерам столбцов с ИНН')

    args = parser.parse_args()
    return args.filename[0], args.outputFileName, args.inn, args.kpp


def main():
    filename, output_file, inns, kpps = argument_parse()
    print(filename, output_file, inns, kpps)
    formatter = VATFormatter.VATFormatter(filename, inns, kpps, output_file)
    formatter.correct_type_of_vat()


if __name__ == '__main__':
    main()
