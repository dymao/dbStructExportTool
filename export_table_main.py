from settings import Settings
import function as fn
import worddoc_util as doc_util


def run_export_main():
    """导出数据库的主方法"""
    settings = Settings()
    # 查询数据库的所有表格
    table_list = fn.query_tables(settings)
    # 查询数据库的所有表格的所有字段信息
    table_info_list = fn.query_table_info(settings, table_list)
    # 创建word文档
    document = doc_util.create_doc(settings)
    # 根据数据，创建word表格
    fn.create_table_list(document, settings, table_info_list)
    # 保存文档到指定路径
    doc_util.save_doc(document, settings.word_doc_file_path)

    # 是否同时生成pdf文档
    if settings.word_to_pdf_flag:
        extension_str = fn.file_extension(settings.word_doc_file_path)
        pdf_path = settings.word_doc_file_path.replace(extension_str, ".pdf")
        doc_util.doc_to_pdf(settings.word_doc_file_path, pdf_path)


if __name__ == '__main__':
    run_export_main()

