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
    parser.add_argument('-i', metavar='INN', dest='inn', action='store', nargs='+', type=int, required=True,
                        help='Номера столбцов, в которых есть ИНН')
    parser.add_argument('-k', metavar='KPP', dest='kpp', action='store', nargs='*', type=int,
                        help='Номера столбцов, в которых есть КПП.\n'
                             'Номера столбцов указывать соответствующие номерам столбцов с ИНН')
    parser.add_argument('-n', metavar='name_column', dest='name', nargs='*', action='store', type=int,
                        help='Номера столбцов, используемых для формирования имени выходных файлов для КА\n'
                             'столбцы будут скопированы и приписаны справа')

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
