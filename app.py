import streamlit as st
import pandas as pd
import os
from io import BytesIO
import xlsxwriter

# 페이지 설정
st.set_page_config(page_title="스마트 견적서 생성기", layout="wide")

# 1. 데이터 로드
@st.cache_data(ttl=10) # 10초마다 엑셀 변경사항 체크 (자주 업데이트하신다고 하여 단축)
def load_data():
    if os.path.exists("products.xlsx"):
        return pd.read_excel("products.xlsx")
    else:
        st.error("'products.xlsx' 파일을 찾을 수 없습니다.")
        return pd.DataFrame()

df = load_data()

st.title("📦 품목 관리 및 견적서 자동 생성")

if not df.empty:
    # 2. 제품 검색 및 관리 화면
    st.subheader("🔍 제품 정보 검색")
    search_col1, search_col2 = st.columns(2)
    with search_col1:
        category = st.selectbox("분류", ["전체"] + list(df['분류'].unique()))
    with search_col2:
        keyword = st.text_input("제품명(국문) 검색")

    filtered_df = df.copy()
    if category != "전체":
        filtered_df = filtered_df[filtered_df['분류'] == category]
    if keyword:
        filtered_df = filtered_df[filtered_df['품명(국문)'].str.contains(keyword, na=False)]

    st.dataframe(filtered_df, use_container_width=True)

    # 3. 견적서 생성 섹션
    st.divider()
    st.subheader("📝 견적서 작성 (오퍼가 수정 가능)")
    
    selected_name = st.selectbox("견적서에 넣을 제품을 선택하세요", filtered_df['품명(국문)'].tolist())
    
    if selected_name:
        item = df[df['품명(국문)'] == selected_name].iloc[0]
        
        # 수정 가능한 오퍼가 입력 칸
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            st.info(f"선택된 제품: {selected_name}")
            img_path = f"images/{item['이미지']}"
            if os.path.exists(img_path):
                st.image(img_path, width=200)
        with col2:
            new_offer_unit = st.number_input("오퍼가 FOB - 단가 수정", value=float(item['오퍼가 FOB -단가']))
        with col3:
            new_offer_ctn = st.number_input("오퍼가 FOB - C/T가격 수정", value=float(item['오퍼가 FOB-C/T가격']))

        # 4. 견적서 엑셀 다운로드 (양식 적용)
        if st.button("📊 견적서 엑셀 파일 생성"):
            output = BytesIO()
            workbook = xlsxwriter.Workbook(output)
            sheet = workbook.add_worksheet("Quotation")

            # 서식 설정
            header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#EFEFEF'})
            data_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            
            # 헤더 정의
            sheet.merge_range('A1:A2', 'PICTURE', header_fmt)
            sheet.merge_range('B1:B2', 'Weight(EA)', header_fmt)
            sheet.merge_range('C1:C2', 'EA/CTN', header_fmt)
            sheet.merge_range('D1:F1', 'Weight, Cbm/ctn', header_fmt) # 상위그룹 묶기
            sheet.write(1, 3, 'net(kg)', header_fmt)
            sheet.write(1, 4, 'gross(kg)', header_fmt)
            sheet.write(1, 5, 'cbm', header_fmt)
            sheet.merge_range('G1:H1', 'FOB KOREAN PORT', header_fmt) # 상위그룹 묶기
            sheet.write(1, 6, 'EA', header_fmt)
            sheet.write(1, 7, 'CTN', header_fmt)
            sheet.merge_range('I1:I2', 'Storage', header_fmt)
            sheet.merge_range('J1:J2', 'Shelf Life', header_fmt)
            sheet.merge_range('K1:K2', 'MOQ', header_fmt)

            # 데이터 행 작성 (3행부터)
            row = 2
            sheet.set_row(row, 100) # 이미지 들어갈 자리 높이 확보
            
            # 1. 이미지 삽입
            if os.path.exists(img_path):
                sheet.insert_image(row, 0, img_path, {'x_scale': 0.15, 'y_scale': 0.15, 'x_offset': 5, 'y_offset': 5})
            else:
                sheet.write(row, 0, "No Image", data_fmt)

            # 2. 나머지 데이터
            sheet.write(row, 1, f"{item['규격(g)']}g", data_fmt)
            sheet.write(row, 2, item['수량/박스'], data_fmt)
            sheet.write(row, 3, item['Weight CBM/CTN - net'], data_fmt)
            sheet.write(row, 4, item['Weight CBM/CTN - gross'], data_fmt)
            sheet.write(row, 5, item['Weight CBM/CTN - CBM'], data_fmt)
            sheet.write(row, 6, new_offer_unit, data_fmt)
            sheet.write(row, 7, new_offer_ctn, data_fmt)
            sheet.write(row, 8, item['storage'], data_fmt)
            sheet.write(row, 9, item['shelf life'], data_fmt)
            sheet.write(row, 10, item['MOQ'], data_fmt)

            sheet.set_column('A:K', 15) # 열 너비 자동 조정 대용

            workbook.close()
            
            st.download_button(
                label="💾 수정된 견적서 다운로드",
                data=output.getvalue(),
                file_name=f"Quotation_{selected_name}.xlsx",
                mime="application/vnd.ms-excel"
            )

else:
    st.warning("데이터가 비어있습니다. 'products.xlsx' 파일을 확인해 주세요.")
