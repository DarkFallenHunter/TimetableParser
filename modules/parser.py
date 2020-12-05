import re


class WeeksParser:
    # Регулярка для поиска недель
    __weeks_reg = r'\(\d[\d,\s]+(\s?н\.?)?\)|\d[\d,\s-]+\s?н\.?'
    # Регулярка для поиска недель кроме заданных
    __except_weeks_reg = r'кр [\d,\s]{1,}\s?н?\.?'

    @classmethod
    def __get_weeks_from_str(cls, weeks_str, parity):
        if type(weeks_str) != str:
            raise TypeError('Weeks must be string')

        weeks = list(map(int, re.findall(r'\d+', weeks_str)))

        if '-' in weeks_str:
            weeks_count = len(weeks)

            if weeks_count <= 1 or weeks_count > 2 or weeks[0] > weeks[1]:
                raise ValueError('Wrong range format - {}'.format(weeks_str))

            weeks = list(range(weeks[0], weeks[1] + 1, 2))

        return [week for week in weeks
                if week in cls.__get_weeks_list_by_parity(parity)]

    @staticmethod
    def __get_weeks_list_by_parity(parity):
        if type(parity) != bool:
            raise TypeError('Parity must be boolean')

        if parity:
            return list(range(2, 17, 2))
        return list(range(1, 17, 2))

    @classmethod
    def __get_weeks_exclude_some(cls, excluded_weeks, parity):
        exc_weeks_type = type(excluded_weeks)

        if exc_weeks_type == str:
            excl_weeks_lst = cls.__get_weeks_from_str(excluded_weeks, parity)
        elif exc_weeks_type == list:
            excl_weeks_lst = excluded_weeks
        else:
            raise TypeError('Excluded weeks can be only in str or list')

        if len(excl_weeks_lst) <= 0:
            raise ValueError('Excluded weeks not expected')

        all_weeks = cls.__get_weeks_list_by_parity(parity)

        return [week for week in all_weeks if week not in excl_weeks_lst]

    @staticmethod
    def __separate_subject_and_weeks(class_name, weeks_match):
        weeks_str = weeks_match[0].strip()
        match_start = weeks_match.start()

        if match_start == 0:
            subject_name = class_name[len(weeks_str):]
        else:
            subject_name = class_name[:match_start]

        return (subject_name.strip(), weeks_str)

    @classmethod
    def parse_subject_and_weeks(cls, class_name, parity):
        except_weeks_match = re.search(cls.__except_weeks_reg, class_name)

        if except_weeks_match:
            subject_name, weeks_str = cls.__separate_subject_and_weeks(
                class_name, except_weeks_match
            )

            weeks = cls.__get_weeks_exclude_some(weeks_str, parity)
        else:
            weeks_match = re.search(cls.__weeks_reg, class_name)

            if weeks_match:
                subject_name, weeks_str = cls.__separate_subject_and_weeks(
                    class_name, weeks_match
                )

                weeks = cls.__get_weeks_from_str(weeks_str, parity)
            else:
                subject_name = class_name
                weeks = cls.__get_weeks_list_by_parity(parity)

        return (subject_name, weeks)


class TimetableParser:
    def __init__(
                 self,
                 merged_cell_class,
                 teachers_list,
                 teacher_col_name='ФИО преподавателя'
                ):
        self.__merged_cell_class = merged_cell_class
        self.__min_row = 4  # Перенести в файл конфига
        self.__teacher_col_name = teacher_col_name
        self.__teachers_list = teachers_list

    @staticmethod
    def __get_str_week_parity(week_str):
        if type(week_str) != str:
            raise TypeError('Week can be only in str')

        if week_str == 'I':
            return False
        elif week_str == 'II':
            return True
        else:
            raise ValueError('Wrong week value, must bu "I" or "II"')

    def __get_teachers_columns(self, sheet):
        return [cell.column - 1 for cell in sheet[3]
                if cell.value == self.__teacher_col_name]

    def __get_week(self, sheet, row_idx):
        start_shift = 0

        while isinstance(
                sheet[row_idx - start_shift][4],
                self.__merged_cell_class
        ):
            start_shift += 1

        return sheet[row_idx - start_shift][4].value

    def __get_class_number(self, sheet, row_idx):
        start_shift = 0

        while isinstance(
                sheet[row_idx - start_shift][1],
                self.__merged_cell_class
        ):
            start_shift += 1

        return sheet[row_idx - start_shift][1].value

    def __find_teacher(self, teacher_name):
        try:
            return next(filter(
                    lambda teacher: teacher_name is not None
                        and teacher in teacher_name,
                        self.__teachers_list))
        except StopIteration:
            return None

    @staticmethod
    def __add_to_result(result: dict, teacher: str, key: tuple, value: dict):
        if teacher not in result:
            result[teacher] = {}

        if key not in result[teacher]:
            result[teacher][key] = []

        for class_info in result[teacher][key]:
            if value['group'] == class_info['group'] and \
               value['clsname'] == class_info['clsname'] and \
               value['clstype'] == class_info['clstype'] and \
               value['clsroom'] == class_info['clsroom']:
                class_info['weeks'] = list(
                    set(class_info['weeks'] + value['weeks'])
                )
                break
        else:
            result[teacher][key].append(value)

    def __add_few_to_result(self, result, teacher, key, values):
        teachers = re.split(r'[+\n\t]| {2,}', teacher)

        if len(teachers) > 1:
            for idx, teacher in teachers:
                self.__add_to_result(
                    result,
                    teacher,
                    key,
                    values[idx]
                )
        else:
            for value in values:
                self.__add_to_result(
                    result,
                    teacher,
                    key,
                    value
                )

    @staticmethod
    def __add_fld_to_values(values, fld_name, fld_val):
        fld_vals = re.split(r'[+\n\t]| {2,}', fld_val)

        if len(fld_vals) > 1:
            for idx, value in enumerate(values):
                value[fld_name] = fld_vals[idx]
        else:
            for value in values:
                value[fld_name] = fld_val

    @classmethod
    def __parse_few_cls_names(cls, cls_names, class_type, classroom, week):
        values = []

        for cls_name in cls_names:
            week_parity = cls.__get_str_week_parity(week)
            clear_cls_name, weeks = WeeksParser.parse_subject_and_weeks(
                cls_name, week_parity
            )
            values.append({
                'weeks': weeks,
                'clsname': clear_cls_name
            })

        cls.__add_fld_to_values(values, 'clstype', class_type)
        cls.__add_fld_to_values(values, 'clsroom', classroom)

        return values

    def parse_sheet(self, sheet):
        result = {}

        for column_idx in self.__get_teachers_columns(sheet):
            week_day = None
            for row in sheet.iter_rows(min_row=self.__min_row):
                row_idx = row[0].row
                founded = False

                if not isinstance(row[0], self.__merged_cell_class):
                    week_day = row[0].value

                if isinstance(row[column_idx], self.__merged_cell_class) \
                   and row[4].value == 'II':
                    week = self.__get_week(sheet, row_idx)
                    shift_row_idx = row_idx - 1

                    while isinstance(
                        sheet[shift_row_idx][column_idx],
                        self.__merged_cell_class
                    ):
                        shift_row_idx -= 1

                    teacher = self.__find_teacher(
                        sheet[shift_row_idx][column_idx].value
                    )

                    if week != self.__get_week(sheet, shift_row_idx) \
                       and teacher is not None:
                        shift_row = sheet[shift_row_idx]
                        founded = True

                else:
                    teacher = self.__find_teacher(row[column_idx].value)

                    if teacher is None:
                        continue

                    shift_row = row
                    week = self.__get_week(sheet, row_idx)
                    founded = True

                if not founded:
                    continue

                # Наименование группы
                group = sheet[2][column_idx - 2].value.strip()
                # Наименование предмета
                class_name = shift_row[column_idx - 2].value.strip()
                # Вид занятий1
                class_type = shift_row[column_idx - 1].value.strip()
                # Аудитория
                classroom = shift_row[column_idx + 1].value.strip()
                # Номер пары
                class_number = self.__get_class_number(sheet, row_idx)

                cls_names = re.split(r'[+\n\t]| {2,}', class_name)

                key = (week_day, class_number)

                if len(cls_names) > 1:
                    values = self.__parse_few_cls_names(
                        cls_names,
                        class_type,
                        classroom,
                        week
                    )

                    for value in values:
                        value['group'] = group

                    self.__add_few_to_result(result, teacher, key, values)
                else:
                    week_parity = self.__get_str_week_parity(week)
                    clear_class_name, weeks = \
                        WeeksParser.parse_subject_and_weeks(
                            class_name, week_parity
                        )
                    value = {
                        'weeks': weeks,
                        'group': group,
                        'clsname': clear_class_name,
                        'clstype': class_type,
                        'clsroom': classroom
                    }

                    self.__add_to_result(result, teacher, key, value)

        return result
