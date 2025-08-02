import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Page

def create_stock_chart(df, chart_title, gold_price_col, stock_col):
    """创建存量值图表，金价范围限定在3200至3400"""
    # 数据处理：处理带逗号的数值并筛选金价范围
    # 1. 处理金价列
    df[gold_price_col] = pd.to_numeric(df[gold_price_col], errors='coerce')
    
    # 2. 处理存量值列（带逗号的情况）
    df[stock_col] = df[stock_col].astype(str).str.replace(',', '', regex=False)
    df[stock_col] = pd.to_numeric(df[stock_col], errors='coerce')
    
    # 3. 删除空值并筛选金价范围为3200至3400
    df = df.dropna(subset=[gold_price_col, stock_col])
    filtered_df = df[(df[gold_price_col] >= 3150) & (df[gold_price_col] <= 3500)]  # 核心修改
    
    # 按金价排序
    filtered_df = filtered_df.sort_values(by=gold_price_col)
    
    # 提取数据
    x_labels = filtered_df[gold_price_col].astype(int).astype(str).tolist()
    stock_values = filtered_df[stock_col].tolist()
    
    # 创建图表
    bar = (
        Bar(init_opts=opts.InitOpts(
            width="800px", 
            height="450px"
        ))
        .add_xaxis(xaxis_data=x_labels)
        .add_yaxis(
            series_name="存量值",
            y_axis=stock_values,
            itemstyle_opts=opts.ItemStyleOpts(color="#4CAF50"),
            label_opts=opts.LabelOpts(
                is_show=True,
                position="top",
                formatter=lambda params: f"{stock_values[params.data_index]:.2f}",
                font_size=9,
                color="#2E7D32"
            )
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(
                title=chart_title,
                pos_left="center",
                title_textstyle_opts=opts.TextStyleOpts(font_size=14)
            ),
            xaxis_opts=opts.AxisOpts(
                name="金价（3200-3400）",  # 标题注明范围
                axislabel_opts=opts.LabelOpts(
                    rotate=45,
                    font_size=10
                ),
                interval=5
            ),
            yaxis_opts=opts.AxisOpts(
                name="存量值",
                split_number=10,
                axislabel_opts=opts.TextStyleOpts(font_size=10)
            ),
            tooltip_opts=opts.TooltipOpts(
                formatter=lambda params: f"金价: {x_labels[params.data_index]}<br>存量值: {params.value:.2f}"
            ),
            legend_opts=opts.LegendOpts(is_show=False)
        )
    )
    return bar

# 配置参数
excel_file = "VoiDetailsForProduct 9.xls"  # 替换为您的Excel文件路径
gold_price_column = "Strike"   # 金价列名
stock_column = "At Close"      # 存量值列名

# 读取数据时处理带逗号的数值
call_df = pd.read_excel(
    excel_file, 
    sheet_name="call",
    converters={stock_column: lambda x: float(str(x).replace(',', '')) if pd.notna(x) else None}
)
put_df = pd.read_excel(
    excel_file, 
    sheet_name="put",
    converters={stock_column: lambda x: float(str(x).replace(',', '')) if pd.notna(x) else None}
)

# 生成存量值图表（金价范围3200-3400）
call_stock_chart = create_stock_chart(call_df, "看涨期权存量值", gold_price_column, stock_column)
put_stock_chart = create_stock_chart(put_df, "看跌期权存量值", gold_price_column, stock_column)

# 确保上下排列
page = Page(layout=Page.SimplePageLayout)
page.add_js_funcs("""
    window.onload = function() {
        var containers = document.getElementsByClassName('chart-container');
        for (var i = 0; i < containers.length; i++) {
            containers[i].style.display = 'block';
            containers[i].style.width = '100%';
            containers[i].style.clear = 'both';
            containers[i].style.margin = '0 auto 30px';
        }
    }
""")
page.add(call_stock_chart, put_stock_chart)
page.render("存量数据.html")
print("已生成金价范围3200-3400的存量值图表：存量数据.html")
    