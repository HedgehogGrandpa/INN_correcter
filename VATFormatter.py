import xlrd
import xlsxwriter
import copy
import os.path as path
import sys


class VATFormatter:
    def __init__(self, ifn, inns, kpps=None, ofn='', prefixs=[], postfix=[]):
        def generate_output_filename(out_fn):
            out_fn.translate(str.maketrans('', '', '?/\\<>|*:"'))
            if out_fn:
                if len(out_fn) > 5 and out_fn.endswith('.xlsx'):
                    return out_fn
                else:
                    return '{}.xlsx'.format(out_fn)
            else:
                return 'Corrected {}'.format(self._input_file_name)

        def generate_inn_kpp_dict(inn_list, kpp_list):
            if kpp_list is None:
                kpp_list = [None] * len(inn_list)
            else:
                kpp_list.extend([None] * (len(inn_list) - len(kpp_list)))
            inn_kpp = {i: k for i, k in zip(inn_list, kpp_list)}
            return inn_kpp

        def check_input_file(fn):
            if path.isfile(fn):
                return fn
            else:
                raise AttributeError('Input file {} does not exist'.format(fn))

        self._cur_row_num = 0
        self._cur_in_row = None
        self._cur_out_row = None
        self._input_file_name = check_input_file(ifn)
        self._output_file_name = generate_output_filename(ofn)
        self._inn_kpp = generate_inn_kpp_dict(inns, kpps)
        self._sheet = xlrd.open_workbook(self._input_file_name).sheet_by_index(0)
        self._work_book = xlsxwriter.Workbook(self._output_file_name)
        self._outsheet = self._work_book.add_worksheet('Corrected')

    # проверка корректности ИНН
    @staticmethod
    def check_inn(inn):
        if len(inn) not in (10, 12):
            return False

        def inn_csum(inn_str):
            k = (3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8)
            pairs = zip(k[11 - len(inn_str):], [int(x) for x in inn_str])
            res = str(sum([k * v for k, v in pairs]) % 11 % 10)
            return res

        if len(inn) == 10:
            return inn[-1] == inn_csum(inn[:-1])
        else:
            return inn[-2:] == inn_csum(inn[:-2]) + inn_csum(inn[:-1])

    # сформировать новые инн и кпп
    def reformat_cells_kpp_info(self, inn, kpp):
        kpp_value = self._cur_in_row[kpp].value
        inn_value = self._cur_in_row[inn].value

        if kpp_value:
            new_inn = '{:0>10}'.format(int(inn_value))
            new_kpp = '{:0>9}'.format(int(kpp_value))
        else:
            new_inn = '{:0>12}'.format(int(inn_value))
            new_kpp = ''
            if new_inn.startswith('00'):
                new_inn = new_inn[2:]

        return new_inn, new_kpp

    # сформировать только инн. Не знаем где искать кпп
    def _reformat_cells_kpp_none(self, inn):
        inn_value = str(int(self._cur_in_row[inn].value))

        l = len(inn_value)
        if l in (9, 11):
            return '0{}'.format(inn_value), ''
        else:
            return inn_value, ''

    # изменить типы и значения по всей таблице
    def correct_type_of_vat(self):
        try:
            for self._cur_in_row in self._sheet.get_rows():
                try:
                    self.correct_row()
                    self.write_corected_row(False)
                except ValueError as e:
                    print('{}'.format(e))
                    self.write_corected_row(True)
        except Exception as e:
            raise Exception('can\'t format {} row {}: {}'.format(self._cur_row_num, self._cur_in_row, e))
        finally:
            self._work_book.close()

    # изменить тип и значение
    def _change_cell_value(self, inn, kpp, new_inn='', new_kpp=''):
        self._cur_out_row[inn].ctype = 1
        self._cur_out_row[inn].value = new_inn
        try:
            self._cur_out_row[kpp].ctype = 1
            self._cur_out_row[kpp].value = new_kpp
        except KeyError:
            pass
        except TypeError:
            pass

    # подправить одну строку
    def correct_row(self):
        self._cur_out_row = copy.deepcopy(self._cur_in_row)
        for inn, kpp in self._inn_kpp.items():
            new_inn = self._cur_in_row[inn].value
            new_kpp = self._cur_in_row[kpp].value if kpp else ''
            try:
                if kpp:
                    new_inn, new_kpp = self.reformat_cells_kpp_info(inn, kpp)
                else:
                    new_inn, new_kpp = self._reformat_cells_kpp_none(inn)
                if not self.check_inn(new_inn):
                    raise ValueError('Wrong VAT in {} row'.format(self._cur_row_num + 1))
            except ValueError as e:
                raise ValueError('can\'t format {} row {}: {}'.format(self._cur_row_num, self._cur_in_row, e))
            finally:
                self._change_cell_value(inn, kpp, new_inn, new_kpp)

    # добавить изменённую строку в результат. Пока всё преобразуется к строковому типу без форматирования
    def write_corected_row(self, error):
        start_cell = 'A{}'.format(self._cur_row_num + 1)
        row_values = [str(cell.value) for cell in self._cur_out_row]
        row_format = self._work_book.add_format({'bold': True, 'font_color': 'red'}) \
            if error else self._work_book.add_format({'bold': False, 'font_color': 'black'})
        self._outsheet.set_column(0, len(row_values), 15)
        self._outsheet.write_row(start_cell, row_values, row_format)
        self._cur_row_num += 1

    def add_prefix_to_column(self, col_num, prefix):
        self._cur_out_row[col_num].ctype = 1
        self._cur_out_row[col_num].value = '{}{}'.format(prefix, self._cur_out_row[col_num].value)

    def add_postfix_to_column(self, col_num, postfix):
        self._cur_out_row[col_num].ctype = 1
        self._cur_out_row[col_num].value = '{}{}'.format(self._cur_out_row[col_num].value, postfix)

    def copy_column_without_spec(self, col_from):
        new_cell = copy.deepcopy(self._cur_out_row[col_from])
        new_cell.ctype = 1
        new_cell.value.translate(str.maketrans('', '', '?/\\<>|*:"'))
        self._cur_in_row.append(new_cell)
