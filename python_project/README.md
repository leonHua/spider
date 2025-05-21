# Python 项目合集

## 项目总体描述
本代码仓库包含多个 Python 模块，每个模块都为特定功能或任务而设计。旨在提供一个用 Python 开发的实用脚本和工具集合。

## 模块

### 可用模块

#### 豆瓣读书爬虫

*   **描述:** 此模块用于爬取豆瓣读书 (book.douban.com) TOP250 图书的数据。
*   **功能:** 它获取每本书的详细信息，包括书名、作者、评分、评价人数、出版年份、出版社、简短引言/描述以及指向该书豆瓣页面的直接链接。
*   **输出:** 爬取的数据保存在名为 `douban_top250.xlsx` 的 Excel 文件中，该文件位于主 `python_project` 目录下。
*   **如何运行:**
    1.  确保您已安装 Python 以及必要的库：`requests`、`beautifulsoup4` 和 `openpyxl`。如果尚未安装，可以使用 pip 进行安装：
        ```bash
        pip install requests beautifulsoup4 openpyxl
        ```
    2.  导航到爬虫模块的目录：
        ```bash
        cd python_project/douban_scraper/
        ```
    3.  运行爬虫脚本：
        ```bash
        python scraper.py
        ```
        脚本将打印其进度，并在数据保存后通知您。

## 未来开发
本项目正在持续开发中。未来将添加更多 Python 模块和实用脚本。敬请期待！
