import db_util_mysql
import worddoc_util as doc_util
import os.path


def get_conn(settings):
    """获取数据库的连接"""
    return db_util_mysql.conn_mysql(settings.mysql_host,
                                    settings.mysql_port,
                                    settings.mysql_user,
                                    settings.mysql_password,
                                    settings.mysql_dbname,
                                    settings.mysql_charset)


def query_tables(settings):
    """获取指定数据库的所有表名 格式： [(表1名, 表1描述), (表2名, 表2描述), ...]"""
    conn = get_conn(settings)
    query_para = [settings.mysql_dbname, 'base table']
    table_tuple = db_util_mysql.select_list(conn, settings.query_table_list_sql, query_para)
    db_util_mysql.close_mysql(conn)
    table_list = []
    if table_tuple:
        for table_name in table_tuple:
            table_list.append(table_name)
    return table_list


def query_table_info(settings, table_list):
    """根据表格列表获取所有表的表结构信息
        存储结构为：List[ {"table1": "表名",
                        "table1_name": "表描述",
                        "table1_field_list": [
                             field1, field2, field3
                           ]
                      }]
    """
    table_info_list = []
    if table_list:
        for table_name in table_list:
            table_info_dict = {"table": table_name[0], "table_name": table_name[1]}
            conn = get_conn(settings)
            query_para = [settings.mysql_dbname, table_name[0]]
            field_info_list = db_util_mysql.select_list(conn, settings.query_table_info_sql, query_para)
            if field_info_list:
                table_field_list = []
                for info_name in field_info_list:
                    field_list = list(info_name[:4])
                    len_str = info_name[4] if info_name[4] else str(info_name[5]) + "," + str(info_name[6]) if info_name[5] else ''
                    field_list.append(len_str)
                    field_list.append(settings.mysql_nullable_dic.get(info_name[7], ''))
                    field_list.append(info_name[8] if info_name[8] else '')
                    col_key = settings.mysql_col_key_dic.get(info_name[9], '')
                    auto_increment = settings.mysql_col_key_dic.get(info_name[10], '')
                    field_list.append(col_key + auto_increment)
                    table_field_list.append(field_list)
                table_info_dict['table_field_list'] = table_field_list
            db_util_mysql.close_mysql(conn)
            table_info_list.append(table_info_dict)
    return table_info_list


def create_table_list(document, settings, table_info_list):
    """循环创建表格"""
    num = 1
    for table_info in table_info_list:
        data = table_info['table_field_list']
        row_nums = len(data)  # 表格行数
        col_nums = len(data[0])  # 表格列数
        # 添加表格段落标题
        graph_title = str(num) + ". " + table_info['table_name'] + "(" + table_info['table'] + ")"
        doc_util.create_paragraph_title(document, graph_title,  settings)
        # 添加表格段落说明
        graph_text = "表名：" + table_info['table']
        doc_util.create_paragraph(document, graph_text, settings)
        # 为每个表对象创建一个表格
        table = doc_util.create_table(document, row_nums + 1, col_nums, settings.table_style)
        # 填充表格标题单元格
        doc_util.fill_table_header_cells(table, settings)
        # 填充表格内容单元格
        doc_util.fill_table_content_cells(table, data, settings, row_nums, col_nums)
        # 添加一个空的段落
        document.add_paragraph("")
        num += 1


def file_extension(path):
    """获取文件后缀名"""
    return os.path.splitext(path)[1]

