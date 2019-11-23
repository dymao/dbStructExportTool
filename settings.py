from docx.shared import RGBColor


class Settings:
    """存储项目所有的常量的类"""

    def __init__(self):
        # mysql 连接信息
        self.mysql_host = '127.0.0.1'
        self.mysql_port = 3307
        self.mysql_user = 'root'
        self.mysql_password = '123456'
        self.mysql_dbname = 'biaozhunyun'
        self.mysql_charset = 'utf8'

        # 定义mysql 常量字典
        self.mysql_col_key_dic = {'PRI': '主键', 'UNI': '唯一索引', 'MUL': '一般索引', 'auto_increment': ',自动递增'}
        self.mysql_nullable_dic = {'YES': 'Y', 'NO': 'N'}

        # 文档保存路径
        self.word_doc_file_path = r'F:/数据库结构说明文档.docx'
        # 是否同时生成pdf
        self.word_to_pdf_flag = True

        # 定义查询信息
        self.query_table_list_sql = 'select table_name, table_comment  from information_schema.tables where table_schema= %s and table_type= %s'
        self.query_table_info_sql = 'select ordinal_position, column_name, column_comment, data_type, character_maximum_length, numeric_precision, numeric_scale, is_nullable, column_default, column_key, extra from information_schema.columns where table_schema= %s and table_name = %s  order by ordinal_position'

        # 表格相关属性定义
        self.table_header = ['序号', '字段名称', '字段描述', '字段类型', '长度', '允许空', '默认值', '备注']  # 表格标题
        self.table_col_width_dic = {0: 1, 1: 1, 2: 3, 3: 1, 4: 1, 5: 1, 6: 1, 7: 1}  # 表格列宽
        self.table_style = 'Table Grid'  # 表格样式
        self.table_font_name = u'微软雅黑'  # 字体样式

        # 段落标题样式
        self.table_title_font_size = 11
        self.table_title_font_level = 3  # 标题级别

        # 定义表格标题样式
        self.table_header_font_bold = True
        self.table_header_font_size = 9
        self.table_header_font_color = RGBColor(0, 0, 0)

        # 定义表格内容样式
        self.table_content_font_bold = False
        self.table_content_font_size = 9
        self.table_content_font_color = RGBColor(0, 0, 0)


