import pymysql


def conn_mysql(mysql_host, mysql_port, mysql_user, mysql_password, mysql_dbname, mysql_charset):
    return pymysql.connect(host=mysql_host,
                           port=mysql_port,
                           user=mysql_user,
                           password=mysql_password,
                           db=mysql_dbname,
                           charset=mysql_charset)


def select_list(conn, sql, data=None):
    """执行操作数据的相关sql"""
    cursor = conn.cursor()
    cursor.execute(sql, data)
    return cursor.fetchall()


def search(conn, sql):
    """执行查询sql"""
    cursor = conn.cursor()
    cursor.execute(sql)
    return cursor.fetchone()


def close_mysql(conn):
    """关闭数据库连接"""
    conn.close()
    conn.cursor().close()

