import streamlit as st
import pandas as pd
import numpy as np
import statsmodels.api as sm
import io
from plotly import graph_objects as go
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.stattools import durbin_watson

# --- 页面配置 ---
st.set_page_config(page_title="FRM 回归分析专业版", layout="wide")

# --- 样式美化 ---
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    .developer-tag { font-size: 14px; color: #6c757d; text-align: right; font-weight: bold; }
    .metric-card { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #007bff; }
    </style>
    """, unsafe_allow_html=True)

# --- 开发者信息 (Sidebar) ---
st.sidebar.markdown("### 🛠️ 系统控制面板")
st.sidebar.info("**DEVELOPER：YANG_XIN_PM**")


# --- 函数：生成模板 ---
def generate_template(mode="SLR"):
    output = io.BytesIO()
    if mode == "SLR":
        # 提供足够样本量以避免 nan
        df = pd.DataFrame({
            "Dependent_Y": [10.5, 12.2, 11.8, 13.1, 14.5, 12.9, 15.2, 16.0],
            "Independent_X1": [2.1, 2.5, 2.3, 2.8, 3.2, 2.7, 3.5, 3.8]
        })
    else:
        df = pd.DataFrame({
            "Dependent_Y": [100, 120, 110, 130, 145, 125, 150, 165, 140, 155],
            "X1_Variable": [10, 12, 11, 13, 14, 12, 15, 16, 14, 15],
            "X2_Variable": [0.5, 0.7, 0.6, 0.8, 0.9, 0.7, 1.0, 1.1, 0.9, 1.0],
            "X3_Variable": [20, 25, 22, 28, 32, 26, 35, 38, 30, 34]
        })
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


# --- 侧边栏：下载模板 ---
st.sidebar.header("1. 下载标准模板")
col1, col2 = st.sidebar.columns(2)
with col1:
    st.download_button("下载SLR模板", data=generate_template("SLR"), file_name="SLR_Template.xlsx")
with col2:
    st.download_button("下载MLR模板", data=generate_template("MLR"), file_name="MLR_Template.xlsx")

# --- 主界面 ---
st.title("📊 FRM 定量分析：线性回归模型工作站")
st.write("集成 Durbin-Watson 自相关检测与 VIF 多重共线性分析。")

# --- 数据上传 ---
uploaded_file = st.file_uploader("上传数据表 (.xls, .xlsx, .csv)", type=["csv", "xls", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("📋 数据预览")
    st.dataframe(df.head(5), use_container_width=True)

    # 参数选择
    st.divider()
    st.subheader("⚙️ 模型配置")
    all_columns = df.columns.tolist()

    col_a, col_b = st.columns(2)
    with col_a:
        y_var = st.selectbox("选择因变量 (Y)", all_columns)
    with col_b:
        x_vars = st.multiselect("选择自变量 (X)", [c for c in all_columns if c != y_var])

    if y_var and x_vars:
        # 自由度检查逻辑
        if len(df) <= len(x_vars) + 1:
            st.error(
                f"❌ 样本量不足！当前样本({len(df)})无法支持包含截距在内的 {len(x_vars) + 1} 个参数估计。请增加数据或减少变量。")
        else:
            # 模型计算
            X = df[x_vars]
            X_with_const = sm.add_constant(X)
            y = df[y_var]
            model = sm.OLS(y, X_with_const).fit()

            # --- 核心指标输出 ---
            st.divider()
            st.subheader("📈 回归统计结果 (Regression Summary)")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("R-Squared", f"{model.rsquared:.4f}")
            c2.metric("Adj. R-Squared", f"{model.rsquared_adj:.4f}")
            c3.metric("F-statistic", f"{model.fvalue:.2f}")

            # 计算 Durbin-Watson
            dw_stat = durbin_watson(model.resid)
            c4.metric("Durbin-Watson", f"{dw_stat:.2f}")

            # DW 指标提示 (FRM 考点)
            if dw_stat < 1.5:
                st.warning("DW 显著低于 2：残差可能存在正自相关。")
            elif dw_stat > 2.5:
                st.warning("DW 显著高于 2：残差可能存在负自相关。")
            else:
                st.success("DW 接近 2：残差序列不相关性良好。")

            # 详细参数表
            st.write("**系数分析表 (Coefficient Table)**")
            summary_df = pd.concat([model.params, model.bse, model.tvalues, model.pvalues], axis=1)
            summary_df.columns = ['Coefficient', 'Std. Error', 't-Stat', 'P-value']
            st.table(summary_df)

            # --- MLR 专属：VIF 检查 ---
            if len(x_vars) > 1:
                st.subheader("🛡️ 多重共线性检测 (VIF Analysis)")
                vif_data = pd.DataFrame()
                vif_data["Variable"] = X.columns
                vif_data["VIF"] = [variance_inflation_factor(X.values, i) for i in range(len(X.columns))]

                col_vif_text, col_vif_table = st.columns([1, 1])
                with col_vif_table:
                    st.dataframe(vif_data, hide_index=True)
                with col_vif_text:
                    if vif_data["VIF"].max() > 10:
                        st.error("警告：发现严重多重共线性 (VIF > 10)！")
                    elif vif_data["VIF"].max() > 5:
                        st.warning("提醒：存在中度多重共线性 (VIF > 5)。")
                    else:
                        st.success("良好：未发现明显的共线性问题。")

            # 可视化 (SLR)
            if len(x_vars) == 1:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=df[x_vars[0]], y=y, mode='markers', name='实际值'))
                fig.add_trace(go.Scatter(x=df[x_vars[0]], y=model.predict(X_with_const), name='拟合直线'))
                fig.update_layout(title="一元线性回归拟合图", xaxis_title=x_vars[0], yaxis_title=y_var)
                st.plotly_chart(fig, use_container_width=True)

            # --- 导出报告 ---
            st.subheader("💾 导出分析报告")
            result_buffer = io.BytesIO()
            with pd.ExcelWriter(result_buffer, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='回归系数')
                if len(x_vars) > 1:
                    vif_data.to_excel(writer, sheet_name='VIF共线性分析', index=False)
                # 导出详细统计量
                stats_df = pd.DataFrame({
                    "Metric": ["R-Squared", "Adj. R-Squared", "F-stat", "Durbin-Watson", "Observations"],
                    "Value": [model.rsquared, model.rsquared_adj, model.fvalue, dw_stat, model.nobs]
                })
                stats_df.to_excel(writer, sheet_name='模型统计量', index=False)

            st.download_button(
                label="下载完整回归参数报告 (.xlsx)",
                data=result_buffer.getvalue(),
                file_name="FRM_Regression_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.warning("👈 请先在左侧下载模板（已更新样本量）并上传您的数据文件。")

# 底部页脚
st.markdown("---")
st.markdown('<p class="developer-tag">DEVELOPER：YANG_XIN_PM | FRM Quantitative Analysis Tool v2.0</p>',
            unsafe_allow_html=True)