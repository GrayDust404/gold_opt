import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Page

def create_variable_chart(df, chart_title, gold_price_col, change_col):
    """创建单个期权变量值图表"""
    # 数据处理
    df[gold_price_col] = pd.to_numeric(df[gold_price_col], errors='coerce')
    df[change_col] = pd.to_numeric(df[change_col], errors='coerce')
    df = df.dropna(subset=[gold_price_col, change_col])
    
    # 筛选金价范围
    filtered_df = df[(df[gold_price_col] >= 3200) & (df[gold_price_col] <= 3700)]
    filtered_df = filtered_df.sort_values(by=gold_price_col)
    
    # 提取数据
    x_labels = filtered_df[gold_price_col].astype(int).astype(str).tolist()
    y_values = filtered_df[change_col].tolist()
    
    # 计算Y轴范围
    max_val = max(y_values) if y_values else 0
    min_val = min(y_values) if y_values else 0
    y_max = max_val * 1.1 if max_val > 0 else 0
    y_min = min_val * 1.1 if min_val < 0 else 0
    
    # 创建图表（移除extra_html_text_after参数）
    bar = (
        Bar(init_opts=opts.InitOpts(
            width="800px", 
            height="450px"
        ))
        .add_xaxis(xaxis_data=x_labels)
        .add_yaxis(
            series_name="变量值",
            y_axis=y_values,
            itemstyle_opts=opts.ItemStyleOpts(
                color=lambda params: "#2196F3" if y_values[params.data_index] >= 0 else "#F44336"
            ),
            label_opts=opts.LabelOpts(
                is_show=True,
                position="top",
                formatter=lambda params: f"{y_values[params.data_index]:.2f}",
                font_size=9
            )
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(
                title=chart_title,
                pos_left="center",
                title_textstyle_opts=opts.TextStyleOpts(font_size=14)
            ),
            xaxis_opts=opts.AxisOpts(
                name="金价",
                axislabel_opts=opts.LabelOpts(
                    rotate=45,
                    font_size=10
                ),
                interval=5
            ),
            yaxis_opts=opts.AxisOpts(
                name="变量值",
                min_=y_min,
                max_=y_max,
                split_number=10,
                axislabel_opts=opts.TextStyleOpts(font_size=10)
            ),
            tooltip_opts=opts.TooltipOpts(
                formatter=lambda params: f"金价: {x_labels[params.data_index]}<br>变量值: {params.value:.2f}"
            ),
            legend_opts=opts.LegendOpts(is_show=False)
        )
    )
    return bar

# 配置参数
excel_file = "VoiDetailsForProduct 9.xls"
gold_price_column = "Strike"
change_column = "Change"

# 读取数据
call_df = pd.read_excel(excel_file, sheet_name="call")
put_df = pd.read_excel(excel_file, sheet_name="put")

# 生成图表
call_chart = create_variable_chart(call_df, "看涨期权变量值", gold_price_column, change_column)
put_chart = create_variable_chart(put_df, "看跌期权变量值", gold_price_column, change_column)

# 使用Page的简单布局并设置页面样式确保上下排列
page = Page(layout=Page.SimplePageLayout)
# 添加自定义CSS样式确保上下排列
page.add_js_funcs("""
    window.onload = function() {
        // 获取页面容器
        var container = document.getElementsByClassName('chart-container');
        for (var i = 0; i < container.length; i++) {
            container[i].style.display = 'block';
            container[i].style.width = '100%';
            container[i].style.clear = 'both';
            container[i].style.margin = '0 auto 30px';
        }
    }
""")
page.add(call_chart)
page.add(put_chart)

# 保存图表
page.render("变量数据.html")
print("变量分析图已生成")