# Word模板自动生成PDF



### 写在前边

​		一款Word模板自动生成PDF的工具。在**Word**中设置 `{{key}}`表示关键字`key`，同时相应**Excel**需要包含关键字`key`以及它的值，即可自动按照**Excel**中的信息填充**Word**模板，进而生成相应PDF。**Word** 模板、**Excel** 信息参考以及相应PDF参考 `test_case`目录。
<img src="https://github.com/lervisnh/Auto_WordTemp2PDF/blob/master/figures/ui.jpg" style="zoom: 70%;" />



### 关键字 `key` 说明

1. Word**模板**中设置的关键字最好在**Excel**中都出现，否则会填充为`nan`
2. **Excel**中的关键字可以多于**Word模板**中的关键字
3. **一定要有 `name` 关键字**，用于结果文件命名，否则只会保留一个文件
4. 如若设置关键字 `phone`，**结果文件名**会添加该关键字下信息
5. 关键字 `y_m_d`用于自动填充当天时间，**Word模板**  可不设置该关键字



### IDE跑着玩

1. `pip install -r requirements.txt`
2. `python Auto_WordTemp2PDF.py`




### 生成EXE文件

1. `pip install pyinstaller`
2. `pyinstaller -F -w Auto_WordTemp2PDF.py`