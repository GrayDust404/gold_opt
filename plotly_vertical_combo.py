import pandas as pd
import plotly.graph_objects as go

FUTURE_SPOT_DIFF = 33

def plot_option(sheet_name, title, output_html):
    df = pd.read_excel("VoiDetailsForProduct.xls", sheet_name=sheet_name)
    df["现货价"] = pd.to_numeric(df["Strike"], errors="coerce") - FUTURE_SPOT_DIFF
    df["变量值"] = pd.to_numeric(df["Change"], errors="coerce")
    df["存量值"] = pd.to_numeric(df["At Close"].astype(str).str.replace(",", ""), errors="coerce")
    df = df.dropna(subset=["现货价", "变量值", "存量值"])
    df = df[(df["现货价"] >= 3200) & (df["现货价"] <= 4000)].sort_values("现货价")

    # 让价格轴不省略
    price_ticks = df["现货价"].astype(int).tolist()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=df["现货价"],
        x=df["变量值"],
        orientation="h",
        name="变量值",
        marker_color=df["变量值"].apply(lambda v: "#2196F3" if v >= 0 else "#F44336")
    ))
    fig.add_trace(go.Scatter(
        y=df["现货价"],
        x=df["存量值"],
        mode="lines+markers",
        name="存量值",
        line=dict(color="#FF9800", width=2)
    ))
    fig.update_layout(
        title=title,
        xaxis=dict(title="数值", side="bottom"),
        yaxis=dict(
            title="现货价（3200-4000）",
            tickmode="array",
            tickvals=price_ticks,
            ticktext=[str(p) for p in price_ticks],
            tickfont=dict(size=10),  # 字体大一点
        ),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        width=700,
        height=2000,  # 增大高度
        margin=dict(l=100, r=40, t=80, b=80),  # 增大上下左右边距
    )
    fig.write_html(output_html)
    print(f"plotly交互图已生成：{output_html}")

plot_option("call", "看涨期权 存量-变量同图（纵向，plotly）", "plotly_纵向存量变量同图_call.html")
plot_option("put", "看跌期权 存量-变量同图（纵向，plotly）", "plotly_纵向存量变量同图_put.html")
