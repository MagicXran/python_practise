import oracledb

connection = oracledb.connect(user='scc', password='scc', service_name='nercar',encoding='gbk')

# 创建游标对象
cursor = connection.cursor()

# 执行SQL查询
cursor.execute(r'SELECT * FROM 总成绩单')

# 遍历结果
for row in cursor:
    # 访问每一行的数据
    # 可以通过索引或字段名访问特定列的值
    column1_value = row[0]
    column2_value = row[1]
    # ...

    # 打印或处理数据
    print(column1_value, column2_value)

# 关闭游标和连接
cursor.close()
connection.close()
