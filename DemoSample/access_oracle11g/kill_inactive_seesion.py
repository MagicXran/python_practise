import cx_Oracle

# # 定义查询SQL语句
# sql = "SELECT sid, serial#, status, last_call_et FROM v$session WHERE username = 'your_username'"
#
# while True:
#     # 执行查询
#     cursor = conn.cursor()
#     cursor.execute(sql)
#
#     # 遍历查询结果
#     for sid, serial # , status, last_call_et in cursor:
#         if status == 'INACTIVE' and last_call_et > 5:
#             # 如果session为inactive且超过5秒，则释放该session
#             print('Session {0}:{1} is inactive for {2} seconds, releasing...'.format(sid, serial
#             # , last_call_et))
#             cursor.execute('ALTER SYSTEM KILL SESSION \'{0},{1}\' IMMEDIATE'.format(sid, serial
#             # ))
#             conn.commit()
#
#             # 关闭游标
#             cursor.close()
#
#             # 等待一段时间后再次执行查询
#             time.sleep(1)
if __name__ == '__main__':
    dsn = cx_Oracle.makedsn("localhost", 1521, service_name="nercar")
    # 连接Oracle数据库
    connection = cx_Oracle.connect(user="scc", password='scc', dsn=dsn,
                                   encoding="UTF-8")
    print(dsn)
    cursor = connection.cursor()
    sql = "SELECT sid, serial#, status, last_call_et FROM v$session WHERE username = 'SCC'"
    cursor.execute(sql)
    rows = cursor.fetchall()
    for row in rows:
        print(row)
