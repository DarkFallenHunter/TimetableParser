import psycopg2 as pg
from psycopg2 import sql, errors


class DbConnectException(Exception):
    '''Класс-исключенеи для обработки ошибки соединения'''

    def __init__(self, username, dbname, *args):
        self.username = username
        self.dbname = dbname
        self.args = args


class DbInsertingException(Exception):
    '''Класс-исключение для обработки ошибки при вставке'''


class DbTransactionException(Exception):
    '''Класс-исключение для обработки ошибки в транзакции'''


class DbPrivilegeException(Exception):
    '''Класс-исключение для обработки ошибки привилегий'''

    def __init__(self, username, *args):
        self.args = args
        self.username = username


class DbSyntaxException(Exception):
    '''Класс-исключение для обработки синтаксических ошибок БД'''


class TimetableDb:
    '''Класс для работы с БД расписания'''

    def __init__(self, params: dict):
        self.__dbname = params['dbname']
        self.__username = params['user']
        self.__password = params['password']
        self.__schema = params['schema']
        self.__cursor = None

    def __insert_class(self, values: dict) -> int:
        '''Запись информации о паре'''

        query = sql.SQL(
            f'INSERT INTO {self.__schema}.v_class ' +
            '("name",number,class_type,week_day,classroom) ' +
            'VALUES (%s, %s, %s, %s, %s)' +
            'RETURNING id'
        )

        try:
            self.__cursor.execute(
                query,
                (
                    values['clsname'],
                    values['clsnumber'],
                    values['clstype'],
                    values['weekday'],
                    values['clsroom']
                )
            )

            return self.__cursor.fetchone()[0]
        except pg.Error:
            raise DbInsertingException(values)

    def __insert_group(self, group: str) -> int:
        '''Добавление новой группы с получением её id'''

        self.__cursor.callproc('dict.get_group_id_by_code', [group, ])

        return self.__cursor.fetchone()[0]

    def __insert_teacher(self, teacher: str) -> int:
        '''Добавление нового преподавателя с получением его id'''

        self.__cursor.callproc('dict.get_teacher_id_by_name', [teacher, ])

        return self.__cursor.fetchone()[0]

    def __insert_weeks(
        self,
        class_id: int,
        weeks: list,
        group_id: int,
        teacher_id: int
    ) -> None:
        '''Запись в БД информации о неделях для предмета'''

        query = sql.SQL(
            f'INSERT INTO {self.__schema}.class_week ' +
            '(class_id,  group_id, teacher_id, week_num) VALUES ' +
            ','.join(['(%s, %s, %s, %s)'] * len(weeks)))

        try:
            args = [(class_id, group_id, teacher_id, week) for week in weeks]
            self.__cursor.execute(
                query,
                [val for sublist in args for val in sublist]
            )
        except pg.Error:
            raise DbInsertingException(class_id, weeks)

    def insert_classes(self, timetable: dict) -> None:
        '''Управление записью расписания в базу данных'''

        try:
            con = pg.connect(
                database=self.__dbname,
                user=self.__username,
                password=self.__password
            )
        except pg.OperationalError:
            raise DbConnectException(self.__dbname, self.__username)

        try:
            self.__cursor = con.cursor()

            for teacher, teacher_clses in timetable.items():
                for cls_key, clses_info in teacher_clses.items():
                    for cls_info in clses_info:
                        class_id = self.__insert_class(
                            {
                                'weekday': cls_key[0],
                                'clsnumber': cls_key[1],
                                'clsname': cls_info['clsname'],
                                'clstype': cls_info['clstype'],
                                'clsroom': cls_info['clsroom']
                            }
                        )

                        group_id = self.__insert_group(cls_info['group'])
                        teacher_id = self.__insert_teacher(teacher)
                        self.__insert_weeks(
                            class_id,
                            cls_info['weeks'],
                            group_id,
                            teacher_id
                        )

            con.commit()
        except errors.InsufficientPrivilege as e:
            con.rollback()
            raise DbPrivilegeException(self.__username, *e.args)
        except errors.SyntaxError as e:
            con.rollback()
            raise DbSyntaxException(*e.args)
        except errors.InFailedSqlTransaction as e:
            con.rollback()
            raise DbTransactionException(*e.args)
        except DbInsertingException:
            con.rollback()
            raise
        finally:
            con.close()
