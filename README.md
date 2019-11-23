# dbStructExportTool
## 数据库表结构导出工具
此工具目前仅支持导出mysql数据表结构说明文档，导出表结构文档可支持word,pdf格式

### 项目运行说明
##### 1. 项目运行需安装如下依赖包：
    * PyMySQL 0.9.3
    * python-doc 0.8.10
    * pywin32 227
##### 2. 配置更改：
    项目配置文件定义在settings中，运行前，需将数据库连接信息更改为你自己的数据库
    表结构导出文档路径及表格样式等可自行在settings中定义
    其中字段：word_to_pdf_flag 表示是否需要同步导出pdf文件

##### 3. 项目运行：
    项目运行主类： export_table_main.py    
      

