# -
根据CME的期权更新数据，对合约中看涨期权和看跌期权的存量数据与变量数据分别生成柱状图
【芝加哥商品交易所黄金期权数据下载链接】
https://www.cmegroup.com/markets/metals/precious/gold.volume.options.html#optionProductId=192
【原始数据处理】
EXCEL中创建两个sheet分别命名为，call和put
以2509合约为例，将原始数据中SEP 25 Calls和SEP 25 Puts的数据分别粘贴到call和put中，保留列名。
【Python生成数据分析图】
分别修改变量.py 存量.py配置参数中excel_file的值为EXCEL的名称如VoiDetailsForProduct 8.1.xls
运行程序即可生成变量和存量的html图。
（环境依赖pyecharts模块没有的需要安装一下）
