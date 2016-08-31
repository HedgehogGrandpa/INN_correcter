# -*- coding: utf-8 -*-
import argparse
import INNFormatter


def argument_parse():
    parser = argparse.ArgumentParser(
        description='Программа проверяет валидность ИНН-КПП, если возможно, то правит '
                    '(пока только исправление ведущих нулей)')
    parser.add_argument('filename', nargs=1,
                        help='Имя файла .xls или .xlsx формата с информацией об организациях')
    parser.add_argument('outputFileName', nargs='?', default='',
                        help='имя выходного файла')
    parser.add_argument('-i', metavar='INN_correcter', dest='inn', action='store', nargs='+', type=int, required=True,
                        help='Номера столбцов, в которых есть ИНН')
    parser.add_argument('-k', metavar='KPP', dest='kpp', action='store', nargs='*', type=int,
                        help='Номера столбцов, в которых есть КПП.\n'
                             'Номера столбцов указывать соответствующие номерам столбцов с ИНН')
    parser.add_argument('-n', metavar='name_column', dest='name', nargs='*', action='store', type=int,
                        help='Номера столбцов, используемых для формирования имени выходных файлов для КА\n'
                             'столбцы будут скопированы и приписаны справа')
    parser.add_argument('-p', metavar='prefix', dest='pre', nargs='*', action='store', type=str,
                        help='Номера столбцов и префиксы для добавления.\n'
                             'Задавать в порядке <Номер_столбца1> <Префикс1> <Номер_столбца2> <Прификс2> ...')
    parser.add_argument('-s', metavar='suffix', dest='suf', nargs='*', action='store', type=str,
                        help='Номера столбцов и суффиксы для добавления.\n'
                             'Задавать в порядке <Номер_столбца1> <Суффикс1> <Номер_столбца2> <Суффикс2> ...')

    args = parser.parse_args()
    return args.filename[0], args.outputFileName, args.inn, args.kpp, args.name, args.pre, args.suf


def correcting(ifn, inns, kpps, ofn, name, pre, suf):
    formatter = INNFormatter.INNFormatter(ifn, inns, kpps, ofn, name, pre, suf)
    formatter.correct_inn()


def gui():
    pass


def main():
    filename, output_file, inns, kpps, name, pre, suf = argument_parse()
    formatter = INNFormatter.INNFormatter(filename, inns, kpps, output_file, name, pre, suf)
    formatter.correct_inn()


if __name__ == '__main__':
    main()
