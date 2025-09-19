import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

FUTURE_SPOT_DIFF = 33

def get_option_fig(df, title):
    df["现货价"] = pd.to_numeric(df["Strike"], errors="coerce") - FUTURE_SPOT_DIFF
    df["变量值"] = pd.to_numeric(df["Change"], errors="coerce")
    df["存量值"] = pd.to_numeric(df["At Close"].astype(str).str.replace(",", ""), errors="coerce")
    df = df.dropna(subset=["现货价", "变量值", "存量值"])
    df = df[(df["现货价"] >= 3200) & (df["现货价"] <= 4000)].sort_values("现货价")
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
            tickfont=dict(size=10),
        ),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        width=700,
        height=900,
        margin=dict(l=100, r=40, t=80, b=80),
    )
    return fig

def calc_max_pain(call_df, put_df):
    """
    计算最大痛苦价值（Max Pain）及其对应的行权价
    """
    strikes = sorted(set(call_df["Strike"]).intersection(set(put_df["Strike"])))
    min_pain = None
    min_strike = None

    for s in strikes:
        # Call持仓损失：max(0, s - 行权价) * 持仓量
        call_loss = ((s - call_df["Strike"]).clip(lower=0) * call_df["存量值"]).sum()
        # Put持仓损失：max(0, 行权价 - s) * 持仓量
        put_loss = ((put_df["Strike"] - s).clip(lower=0) * put_df["存量值"]).sum()
        total_loss = call_loss + put_loss
        if (min_pain is None) or (total_loss < min_pain):
            min_pain = total_loss
            min_strike = s
    return min_strike, min_pain

# 读取数据
call_df = pd.read_excel("VoiDetailsForProduct.xls", sheet_name="call")
put_df = pd.read_excel("VoiDetailsForProduct.xls", sheet_name="put")

# 创建子图
fig = make_subplots(
    rows=1, cols=2,
    shared_yaxes=False,
    subplot_titles=("看涨期权 存量-变量同图（横向，plotly）", "看跌期权 存量-变量同图（横向，plotly）")
)

# call
call_df["现货价"] = pd.to_numeric(call_df["Strike"], errors="coerce") - FUTURE_SPOT_DIFF
call_df["变量值"] = pd.to_numeric(call_df["Change"], errors="coerce")
call_df["存量值"] = pd.to_numeric(call_df["At Close"].astype(str).str.replace(",", ""), errors="coerce")
call_df = call_df.dropna(subset=["现货价", "变量值", "存量值"])
call_df = call_df[(call_df["现货价"] >= 3200) & (call_df["现货价"] <= 4000)].sort_values("现货价")
call_price_ticks = call_df["现货价"].astype(int).tolist()

fig.add_trace(go.Bar(
    y=call_df["现货价"],
    x=call_df["变量值"],
    orientation="h",
    name="变量值(call)",
    marker_color=call_df["变量值"].apply(lambda v: "#2196F3" if v >= 0 else "#F44336")
), row=1, col=1)
fig.add_trace(go.Scatter(
    y=call_df["现货价"],
    x=call_df["存量值"],
    mode="lines+markers",
    name="存量值(call)",
    line=dict(color="#FF9800", width=2)
), row=1, col=1)

# put
put_df["现货价"] = pd.to_numeric(put_df["Strike"], errors="coerce") - FUTURE_SPOT_DIFF
put_df["变量值"] = pd.to_numeric(put_df["Change"], errors="coerce")
put_df["存量值"] = pd.to_numeric(put_df["At Close"].astype(str).str.replace(",", ""), errors="coerce")
put_df = put_df.dropna(subset=["现货价", "变量值", "存量值"])
put_df = put_df[(put_df["现货价"] >= 3200) & (put_df["现货价"] <= 4000)].sort_values("现货价")
put_price_ticks = put_df["现货价"].astype(int).tolist()

fig.add_trace(go.Bar(
    y=put_df["现货价"],
    x=put_df["变量值"],
    orientation="h",
    name="变量值(put)",
    marker_color=put_df["变量值"].apply(lambda v: "#2196F3" if v >= 0 else "#F44336")
), row=1, col=2)
fig.add_trace(go.Scatter(
    y=put_df["现货价"],
    x=put_df["存量值"],
    mode="lines+markers",
    name="存量值(put)",
    line=dict(color="#FF9800", width=2)
), row=1, col=2)

fig.update_yaxes(
    title_text="现货价（3200-4000）",
    tickmode="array",
    tickvals=call_price_ticks,
    ticktext=[str(p) for p in call_price_ticks],
    tickfont=dict(size=10),
    row=1, col=1
)
fig.update_yaxes(
    title_text="现货价（3200-4000）",
    tickmode="array",
    tickvals=put_price_ticks,
    ticktext=[str(p) for p in put_price_ticks],
    tickfont=dict(size=10),
    row=1, col=2
)
fig.update_xaxes(title_text="数值", row=1, col=1)
fig.update_xaxes(title_text="数值", row=1, col=2)

fig.update_layout(
    height=2000,
    width=1400,
    margin=dict(l=100, r=40, t=100, b=80),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    showlegend=True,
    title_text="看涨/看跌期权 存量-变量同图（横向，plotly）"
)

# call_df和put_df已处理好
call_df["Strike"] = pd.to_numeric(call_df["Strike"], errors="coerce")
put_df["Strike"] = pd.to_numeric(put_df["Strike"], errors="coerce")
max_pain_strike, max_pain_value = calc_max_pain(call_df, put_df)
print(f"最大痛苦价值对应行权价: {max_pain_strike}, 最大痛苦价值: {max_pain_value}")

# 在两个子图上添加最大痛苦价值的线
max_pain_y = max_pain_strike - FUTURE_SPOT_DIFF
call_xmin, call_xmax = call_df["变量值"].min(), call_df["变量值"].max()
put_xmin, put_xmax = put_df["变量值"].min(), put_df["变量值"].max()

fig.add_shape(
    type="line",
    x0=call_xmin, x1=call_xmax,
    y0=max_pain_y, y1=max_pain_y,
    xref="x", yref="y",
    line=dict(color="purple", width=2, dash="dash"),
    row=1, col=1
)
fig.add_shape(
    type="line",
    x0=put_xmin, x1=put_xmax,
    y0=max_pain_y, y1=max_pain_y,
    xref="x2", yref="y2",
    line=dict(color="purple", width=2, dash="dash"),
    row=1, col=2
)
fig.add_trace(go.Scatter(
    x=[None], y=[None],
    mode="lines",
    line=dict(color="purple", width=2, dash="dash"),
    name="最大痛苦价值"
), row=1, col=1)

fig.write_html("plotly_横向存量变量同图_all.html")
print("plotly交互图已生成：plotly_横向存量变量同图_all.html")

