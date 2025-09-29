import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import argparse

FUTURE_SPOT_DIFF = 0  # 期货-现货的价差
PRICE_MIN = 2005
PRICE_MAX = 4250

def extract_call_put_df(filepath, contract_name=None):
    # 读取整个VOI Details Report sheet
    df = pd.read_excel(filepath, sheet_name=0, header=None)
    # 找到 American Options 行
    opt_type_idx = df[df[0] == "OPTION TYPE: American Options"].index[0]
    # 合约名在下一行
    main_contract = df.iloc[opt_type_idx + 1, 0]
    # 提取主力合约名（如"OCT 25"）
    auto_contract_prefix = main_contract.split(" Calls")[0].split(" Puts")[0].strip()
    # 使用参数或自动识别
    contract_prefix = contract_name if contract_name else auto_contract_prefix
    call_tag = f"{contract_prefix} Calls"
    put_tag = f"{contract_prefix} Puts"

    # 找到call表头行
    call_header_idx = df[df[0] == call_tag].index
    if len(call_header_idx) == 0:
        raise ValueError(f"未找到合约 {contract_prefix} 的 Calls 数据")
    call_header_idx = call_header_idx[0] + 1
    # 找到call TOTALS行
    call_totals_idx = df[(df[0].astype(str).str.startswith("TOTALS")) & (df.index > call_header_idx)].index[0]
    # 提取call数据（包含表头和TOTALS行）
    call_table = df.iloc[call_header_idx:call_totals_idx+1, :]
    call_table.columns = call_table.iloc[0]
    call_table = call_table[1:]  # 去掉表头行

    # 找到put表头行
    put_header_idx = df[df[0] == put_tag].index
    if len(put_header_idx) == 0:
        raise ValueError(f"未找到合约 {contract_prefix} 的 Puts 数据")
    put_header_idx = put_header_idx[0] + 1
    # 找到put TOTALS行
    put_totals_idx = df[(df[0].astype(str).str.startswith("TOTALS")) & (df.index > put_header_idx)].index[0]
    # 提取put数据（包含表头和TOTALS行）
    put_table = df.iloc[put_header_idx:put_totals_idx+1, :]
    put_table.columns = put_table.iloc[0]
    put_table = put_table[1:]  # 去掉表头行

    # 重置索引
    call_table = call_table.reset_index(drop=True)
    put_table = put_table.reset_index(drop=True)
    return call_table, put_table, contract_prefix

def calc_max_pain(call_df, put_df):
    """
    计算最大痛苦价值（Max Pain）及其对应的行权价
    """
    # 用所有出现过的行权价，不取交集
    strikes = sorted(set(pd.to_numeric(call_df["Strike"], errors="coerce").dropna())
                     .union(set(pd.to_numeric(put_df["Strike"], errors="coerce").dropna())))
    min_pain = None
    min_strike = None

    for s in strikes:
        call_loss = ((s - call_df["Strike"]).clip(lower=0) * call_df["存量值"]).sum()
        put_loss = ((put_df["Strike"] - s).clip(lower=0) * put_df["存量值"]).sum()
        total_loss = call_loss + put_loss
        # print(f"strike: {s}, call_loss: {call_loss}, put_loss: {put_loss}, total_loss: {total_loss}")
        if (min_pain is None) or (total_loss < min_pain):
            min_pain = total_loss
            min_strike = s
    # print(f"min_strike: {min_strike}, min_pain: {min_pain}")
    return min_strike, min_pain

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--contract", type=str, default=None, help="指定合约名（如 'DEC 24'），不指定则自动识别主力合约")
    args = parser.parse_args()

    # 读取数据（自动提取call/put数据和合约名）
    call_df, put_df, contract_name = extract_call_put_df("VoiDetailsForProduct.xls", contract_name=args.contract)

    # call
    call_df["现货价"] = pd.to_numeric(call_df["Strike"], errors="coerce") - FUTURE_SPOT_DIFF
    call_df["变量值"] = pd.to_numeric(call_df["Change"], errors="coerce")
    call_df["存量值"] = pd.to_numeric(call_df["At Close"].astype(str).str.replace(",", ""), errors="coerce")
    call_df = call_df.dropna(subset=["现货价", "变量值", "存量值"])

    # put
    put_df["现货价"] = pd.to_numeric(put_df["Strike"], errors="coerce") - FUTURE_SPOT_DIFF
    put_df["变量值"] = pd.to_numeric(put_df["Change"], errors="coerce")
    put_df["存量值"] = pd.to_numeric(put_df["At Close"].astype(str).str.replace(",", ""), errors="coerce")
    put_df = put_df.dropna(subset=["现货价", "变量值", "存量值"])

    # call_df和put_df已处理好
    call_df["Strike"] = pd.to_numeric(call_df["Strike"], errors="coerce")
    put_df["Strike"] = pd.to_numeric(put_df["Strike"], errors="coerce")
    max_pain_strike, max_pain_value = calc_max_pain(call_df, put_df)
    print(f"最大痛苦价值对应行权价: {max_pain_strike}, 最大痛苦价值: {max_pain_value}")

    # 创建子图
    fig = make_subplots(
        rows=1, cols=2,
        shared_yaxes=False,
        subplot_titles=("看涨期权", "看跌期权")
    )

    call_df = call_df[(call_df["现货价"] >= PRICE_MIN) & (call_df["现货价"] <= PRICE_MAX)].sort_values("现货价")
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

    put_df = put_df[(put_df["现货价"] >= PRICE_MIN) & (put_df["现货价"] <= PRICE_MAX)].sort_values("现货价")
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
        title_text=f"现货价（{PRICE_MIN}-{PRICE_MAX}）",
        tickmode="array",
        tickvals=call_price_ticks,
        ticktext=[str(p) for p in call_price_ticks],
        tickfont=dict(size=10),
        row=1, col=1
    )
    fig.update_yaxes(
        title_text=f"现货价（{PRICE_MIN}-{PRICE_MAX}）",
        tickmode="array",
        tickvals=put_price_ticks,
        ticktext=[str(p) for p in put_price_ticks],
        tickfont=dict(size=10),
        row=1, col=2
    )
    fig.update_xaxes(title_text="数值", row=1, col=1)
    fig.update_xaxes(title_text="数值", row=1, col=2)

    fig.update_layout(
        height=4000,
        width=1400,
        margin=dict(l=100, r=40, t=100, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        showlegend=True,
        title_text=f"{contract_name} 合约概况"
    )

    # 在两个子图上添加最大痛苦价值的线
    max_pain_y = max_pain_strike - FUTURE_SPOT_DIFF
    call_xmin, call_xmax = call_df["存量值"].min(), call_df["存量值"].max()
    put_xmin, put_xmax = put_df["存量值"].min(), put_df["存量值"].max()

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

    fig.write_html("CME Group黄金期权概况.html")
    print("plotly交互图已生成: CME Group黄金期权概况.html")

