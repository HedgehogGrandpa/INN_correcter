import xlrd
import xlsxwriter
import copy


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

        self._input_file_name = ifn
        try:
            self._output_file_name = generate_output_filename(ofn)
            self._inn_kpp = generate_inn_kpp_dict(inns, kpps)
            self._sheet = xlrd.open_workbook(self._input_file_name).sheet_by_index(0)
            self._work_book = xlsxwriter.Workbook(self._output_file_name)
            self._outsheet = self._work_book.add_worksheet('Corrected')
        except Exception as e:
            print(e)
        self._cur_row_num = 0
        self._cur_in_row = None
        self._cur_out_row = None

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

    def _reformat_cells_kpp_none(self, inn):
        inn_value = str(int(self._cur_in_row[inn].value))

        l = len(inn_value)
        if l in (9, 11):
            return '0{}'.format(inn_value)
        else:
            return inn_value

    def correct_type_of_vat(self):
        try:
            for row in self._sheet.get_rows():
                self.correct_row(row)
                self.write_corected_row()
            self._work_book.close()
        except PermissionError as e:
            print('{}'.format(e))

    def change_cell_value(self, inn, kpp, new_inn='', new_kpp=''):
        self._cur_out_row[inn].ctype = 1
        self._cur_out_row[inn].value = new_inn
        try:
            self._cur_out_row[kpp].ctype = 1
            self._cur_out_row[kpp].value = new_kpp
        except KeyError:
            pass
        except TypeError:
            pass

    def correct_row(self, row):
        self._cur_in_row = row
        self._cur_out_row = copy.deepcopy(row)
        for inn in self._inn_kpp:
            try:
                kpp = self._inn_kpp[inn]
                if kpp:
                    try:
                        new_inn, new_kpp = self.reformat_cells_kpp_info(inn, kpp)
                    except Exception as e:
                        self.change_cell_value(inn, kpp)
                        raise Exception('can\'t format row {} in {}, {} columns : {}'
                                        .format(self._cur_in_row, inn, kpp, e))
                else:
                    try:
                        new_inn = self._reformat_cells_kpp_none(inn)
                    except Exception as e:
                        self.change_cell_value(inn, kpp)
                        raise Exception('can\'t format row {} in {} column : {}'.format(self._cur_in_row, inn, e))

                if self.check_inn(new_inn):
                    self.change_cell_value(inn, kpp, new_inn, new_kpp)
                else:
                    self.change_cell_value(inn, kpp)
                    print('Wrong VAT in {} row'.format(self._cur_row_num + 1))
            except Exception as e:
                print('{} in row {}'.format(e, self._cur_row_num + 1))

    def write_corected_row(self):
        self._outsheet.write_row('A{}'.format(self._cur_row_num + 1), [str(cell.value) for cell in self._cur_out_row])
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
