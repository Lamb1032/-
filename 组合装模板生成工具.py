import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="鹿小漫组合装生成工具", layout="wide")
st.title("鹿小漫组合装模板生成工具")

# 初始化 session_state
for key in ['combo_results', 'flag_add_success', 'flag_delete_success', 'flag_clear_success', 'clear_select']:
    if key not in st.session_state:
        st.session_state[key] = False if "flag" in key or "clear" in key else []

# 上传 Excel 文件
file = st.file_uploader("上传包含颜色、尺码、款式、基本售价等信息的原始数据表", type=["xlsx"])

if file:
    df = pd.read_excel(file)
    st.success("原始数据读取成功，请选择商品进行组合")

    show_cols = [c for c in df.columns if any(k in c for k in ["颜色", "尺码", "规格", "款式", "编码", "售价", "名称"])]
    st.dataframe(df[show_cols])

    def format_row(i, row):
        color = row.get("颜色", "")
        size = row.get("规格", row.get("尺码", ""))
        return f"{i} - {row['商品名称']}（{color}/{size}）"

    label_to_index = {format_row(i, row): i for i, row in df.iterrows()}

    selected_labels = st.multiselect(
        "选择要组合的商品",
        options=list(label_to_index.keys()),
        default=[] if st.session_state.clear_select else st.session_state.get("row_select", []),
        key="row_select"
    )
    st.session_state.clear_select = False

    if len(selected_labels) > 5:
        st.warning("最多只能选 5 个商品进行组合")
    elif 2 <= len(selected_labels) <= 5:
        indices = [label_to_index[lbl] for lbl in selected_labels]
        selected_rows = [df.loc[i] for i in indices]

        if st.button("➕ 添加当前组合"):
            names = [row["商品名称"] for row in selected_rows]
            colors = [row.get("颜色", "") for row in selected_rows]
            sizes = [row.get("规格", row.get("尺码", "")) for row in selected_rows]
            codes = [row["商品编码"] for row in selected_rows]
            prices = [row["基本售价"] for row in selected_rows]

            size_combo = "+".join(sizes) if len(set(sizes)) > 1 else sizes[0]
            combo_name = "+".join(names) + "/" + "+".join(colors) + size_combo
            combo_spec = "+".join(names) + "/" + "+".join(colors) + "；" + size_combo

            for i, (code, price) in enumerate(zip(codes, prices)):
                st.session_state.combo_results.append({
                    '组合商品编码': '',
                    '组合装商品标签': '',
                    '组合款式编码': '',
                    '组合商品名称': '',
                    '组合商品简称': combo_name if i == 0 else '',
                    '组合颜色规格': combo_spec if i == 0 else '',
                    '禁止库存同步': '',
                    '商品编码': code,
                    '数量': 1,
                    '应占售价': price,
                    '组合基本售价': '',
                    '组合成本价': '',
                    '是否支持留言备注换货': '',
                    '图片': '',
                    '品牌': ''
                })

            st.session_state.flag_add_success = True
            st.session_state.clear_select = True
            st.rerun()
    elif len(selected_labels) == 1:
        st.warning("组合最少需要两个商品")

    if st.session_state.flag_add_success:
        st.success("当前组合已添加，可以继续选择商品进行组合")
        st.session_state.flag_add_success = False
    if st.session_state.flag_delete_success:
        st.success("已成功删除所选组合")
        st.session_state.flag_delete_success = False
    if st.session_state.flag_clear_success:
        st.success("已成功清空所有组合")
        st.session_state.flag_clear_success = False

    if st.session_state.combo_results:
        st.subheader("所有组合结果（已添加）：")
        df_result = pd.DataFrame(st.session_state.combo_results)
        st.dataframe(df_result)

        combo_names = df_result[df_result["组合商品简称"] != ""]["组合商品简称"].tolist()
        to_delete = st.multiselect("选择要删除的组合", options=combo_names, key="delete_select")

        if st.button("删除所选组合"):
            keep = []
            i = 0
            while i < len(st.session_state.combo_results):
                if st.session_state.combo_results[i]["组合商品简称"] in to_delete:
                    group_name = st.session_state.combo_results[i]["组合商品简称"]
                    while i < len(st.session_state.combo_results) and st.session_state.combo_results[i]["组合商品简称"] == group_name:
                        i += 1
                else:
                    group_name = st.session_state.combo_results[i]["组合商品简称"]
                    while i < len(st.session_state.combo_results) and st.session_state.combo_results[i]["组合商品简称"] == group_name:
                        keep.append(st.session_state.combo_results[i])
                        i += 1
            st.session_state.combo_results = keep
            st.session_state.flag_delete_success = True
            st.rerun()

        if st.button("清空所有组合"):
            st.session_state.combo_results.clear()
            st.session_state.flag_clear_success = True
            st.rerun()

        st.markdown("---")
        start_code = st.text_input("请输入起始组合商品编码（如 Z2000）", value="Z2000")

        if st.button("填充组合商品编码"):
            if start_code[0].isalpha() and start_code[1:].isdigit():
                prefix, num = start_code[0], int(start_code[1:])
                index = 0
                for row in st.session_state.combo_results:
                    if row["组合商品简称"]:
                        row["组合商品编码"] = f"{prefix}{num + index}"
                        index += 1
                    else:
                        row["组合商品编码"] = ""  # 非首行不填编码
                st.success("已成功填充组合商品编码")
                st.rerun()
            else:
                st.error("请输入正确格式的编码，例如 Z2000")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_result.to_excel(writer, index=False, sheet_name="组合装")

        st.download_button(
            "下载组合装 Excel",
            data=output.getvalue(),
            file_name="组合装模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
