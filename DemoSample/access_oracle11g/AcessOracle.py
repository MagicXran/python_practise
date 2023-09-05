import cx_Oracle


class oracleAPI:
    # 构造函数
    def __init__(self, user, pwd, ip, port, sid):
        self.__user = user
        self.__pwd = pwd
        self.ip = ip
        self.__port = port
        self.__sid = sid
        self.__dsn = self.get_dsn()
        self.__db = self.get_conn()
        self.__curs = self.get_curs()

    # 析构函数
    def __del__(self):
        self.__curs.close()
        print('Cursor closed.')
        self.__db.close()
        print('db closed.')

    # 方法
    def get_dsn(self):
        dsn = cx_Oracle.makedsn(self.ip, self.__port, self.__sid)
        return dsn

    def get_conn(self):
        db = cx_Oracle.connect(self.__user, self.__pwd, self.__dsn)
        return db

    def get_curs(self):
        curs = self.__db.cursor()
        return curs

    def get_db_version(self):
        return self.__db.version

    def execute(self, sql):
        result = self.__curs.execute(sql)
        return result


def main():
    user = 'scc'
    pwd = 'scc'
    ip = '10.1.0.158'
    port = '1521'
    sid = '123'
    oapi = oracleAPI(user, pwd, ip, port, sid)
    print(oapi.get_db_version())
    result = oapi.get_curs().execute('select * from dba_users')
    print(result.fetchmany())


if __name__ == "__main__":
    main()
