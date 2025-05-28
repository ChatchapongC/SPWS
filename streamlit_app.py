import streamlit as st
import pandas as pd
import io

st.title("โปรแกรมแปลงไฟล์ Excel สำหรับ SPWS")
st.write("โปรแกรมนี้ใช้สำหรับแปลงไฟล์ Excel ที่ได้จาก SPWS (ส่วนลด 34%) เพื่อเตรียมข้อมูลสำหรับการนำเข้าในระบบ Bigseller by DK")

uploaded_files = st.file_uploader("แนบไฟล์ Excel (.xls, .xlsx) ได้หลายไฟล์พร้อมกัน", type=["xls", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file, skiprows=10, header=None)
        nan_count = 0
        rows_to_write = []
        for _, row in df.iterrows():
            if row.isna().all():
                nan_count += 1
            else:
                nan_count = 0
            if nan_count > 1:
                break
            rows_to_write.append(row)
        df_new = pd.DataFrame(rows_to_write)
        df_new = df_new.dropna(how='all')
        if df_new.shape[1] > 21:
            df_new[22] = df_new[21] - (df_new[21] * 0.34)
        all_data.append(df_new)

    if all_data:
        df_all = pd.concat(all_data, ignore_index=True)
        # ตั้งชื่อคอลัมน์
        col_names = []
        for i in range(df_all.shape[1]):
            if i == 5:
                col_names.append('*ชื่อSKU')
            elif i == 17:
                col_names.append('*จำนวน Stock-In')
            elif i == 22:
                col_names.append('ราคาต่อหน่วย')
            else:
                col_names.append(f'col{i+1}')
        df_all.columns = col_names

        # Consolidate
        df_consolidated = (
            df_all.groupby('*ชื่อSKU', as_index=False)
            .agg({
                '*จำนวน Stock-In': 'sum',
                'ราคาต่อหน่วย': 'first'
            })
        )

        st.write("ผลลัพธ์ Consolidate")
        st.dataframe(df_consolidated)

        # ดาวน์โหลดไฟล์ (แก้ไขให้ใช้ BytesIO)
        output = io.BytesIO()
        df_consolidated.to_excel(output, index=False, header=True)
        output.seek(0)
        st.download_button(
            label="ดาวน์โหลดไฟล์ Excel",
            data=output,
            file_name="SPWS-นำเข้าไฟล์.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )