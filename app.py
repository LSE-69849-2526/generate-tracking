from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from po_to_pi import convert_uploaded_pdfs


st.set_page_config(page_title="PI 工作台", page_icon="PI", layout="wide")
st.markdown(
    """
    <style>
        .block-container {max-width: 1180px; padding-top: 2.25rem; padding-bottom: 3rem;}
        [data-testid="stTabs"] [data-baseweb="tab-list"] {gap: 1.5rem;}
        [data-testid="stTabs"] button {font-size: 1rem; font-weight: 600;}
        [data-testid="stFileUploader"] {padding: 0.5rem 0;}
        div[data-testid="stMetric"] {padding: 0.25rem 0;}
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("PI 工作台")
st.caption("采购订单生成 PI，或将已确认的 PI 汇总为 Tracking。")


def build_tracking_workbook(uploaded_files):
    records = []
    for uploaded_file in uploaded_files:
        workbook = load_workbook(BytesIO(uploaded_file.getvalue()), data_only=True)
        sheet = workbook.active
        pi_number = sheet["F8"].value or ""
        po_number = sheet["F9"].value or ""
        bill_to = str(sheet["B8"].value or "").strip()
        ship_to = str(sheet["B10"].value or "").strip()

        row = 15
        while sheet[f"B{row}"].value:
            try:
                quantity = float(sheet[f"D{row}"].value or 0)
                price = float(sheet[f"E{row}"].value or 0)
            except (TypeError, ValueError):
                quantity = 0.0
                price = 0.0
            records.append(
                {
                    "Item": sheet[f"A{row}"].value or "",
                    "PO number": po_number,
                    "PI number": pi_number,
                    "Billing To": bill_to,
                    "Delivery Address": ship_to,
                    "SKU Number": sheet[f"B{row}"].value,
                    "SKU Name": sheet[f"C{row}"].value or "",
                    "Order Qty": quantity,
                    "Price": price,
                    "Total Amount": quantity * price,
                }
            )
            row += 1

    output = BytesIO()
    pd.DataFrame(records).to_excel(output, index=False)
    return output.getvalue(), len(records)


if "po_result" not in st.session_state:
    st.session_state.po_result = None
if "tracking_result" not in st.session_state:
    st.session_state.tracking_result = None

po_tab, tracking_tab = st.tabs(["PO 转 PI", "PI 转 Tracking"])

with po_tab:
    upload_column, guide_column = st.columns([1.7, 1], gap="large")
    with upload_column:
        st.subheader("上传采购订单")
        uploaded_pdfs = st.file_uploader(
            "选择一个或多个 PO PDF",
            type=["pdf"],
            accept_multiple_files=True,
            key="po_uploads",
        )
        if st.button("生成 PI 文件包", type="primary", disabled=not uploaded_pdfs):
            with st.spinner("正在生成 PI..."):
                zip_data, results = convert_uploaded_pdfs(uploaded_pdfs)
            st.session_state.po_result = {"zip_data": zip_data, "results": results}

    with guide_column:
        st.subheader("操作步骤")
        st.markdown("1. 上传 PO PDF。\n2. 生成并下载 ZIP。\n3. 确认 PI 后，再汇总 Tracking。")
        st.caption("每份 PO 会与对应 PI 一起放入 ZIP 子文件夹。")

    if st.session_state.po_result:
        results = st.session_state.po_result["results"]
        successes = [item for item in results if item["status"] == "success"]
        failures = [item for item in results if item["status"] == "error"]
        review_items = [
            {"来源文件": item["source_name"], "需核对产品": product}
            for item in successes
            for product in item["review_items"]
        ]
        st.divider()
        st.subheader("处理结果")
        metric_one, metric_two, metric_three = st.columns(3)
        metric_one.metric("成功生成", len(successes))
        metric_two.metric("需要核对", len(review_items))
        metric_three.metric("处理失败", len(failures))
        if successes:
            st.download_button(
                "下载 PO 与 PI 文件包",
                data=st.session_state.po_result["zip_data"],
                file_name=f"PO_PI_package_{datetime.now():%Y%m%d_%H%M%S}.zip",
                mime="application/zip",
                type="primary",
            )
            st.dataframe(
                pd.DataFrame(
                    [
                        {"PO 编号": item["po_number"], "PI 编号": item["pi_number"], "来源文件": item["source_name"]}
                        for item in successes
                    ]
                ),
                hide_index=True,
                use_container_width=True,
            )
        if review_items:
            st.warning("以下产品未在 SKU 表中完整匹配，请确认后再使用。")
            st.dataframe(pd.DataFrame(review_items), hide_index=True, use_container_width=True)
        if failures:
            st.error("以下文件未能完成处理。")
            st.dataframe(pd.DataFrame(failures), hide_index=True, use_container_width=True)

with tracking_tab:
    upload_column, guide_column = st.columns([1.7, 1], gap="large")
    with upload_column:
        st.subheader("上传已确认的 PI")
        uploaded_pis = st.file_uploader(
            "选择一个或多个 PI Excel 文件",
            type=["xlsx"],
            accept_multiple_files=True,
            key="pi_uploads",
        )
        if st.button("生成 Tracking 汇总表", type="primary", disabled=not uploaded_pis):
            try:
                tracking_data, row_count = build_tracking_workbook(uploaded_pis)
                st.session_state.tracking_result = {"data": tracking_data, "row_count": row_count}
            except Exception as exc:
                st.session_state.tracking_result = {"error": str(exc)}

    with guide_column:
        st.subheader("操作步骤")
        st.markdown("1. 上传已检查的 PI。\n2. 生成 Tracking 汇总表。\n3. 下载 Excel 文件。")
        st.caption("地址会保留 PI 中原有的多行格式。")

    if st.session_state.tracking_result:
        result = st.session_state.tracking_result
        st.divider()
        if result.get("error"):
            st.error(f"Tracking 生成失败：{result['error']}")
        elif result["row_count"]:
            st.success(f"已生成 {result['row_count']} 条 Tracking 明细。")
            st.download_button(
                "下载 Tracking 汇总表",
                data=result["data"],
                file_name=f"Batch_Tracking_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        else:
            st.warning("上传的 PI 中没有可汇总的产品行。")
