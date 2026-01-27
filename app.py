import streamlit as st
import pandas as pd
import os
from io import BytesIO
from PIL import Image
import xlsxwriter

# 페이지 설정
st.set_page_config(page_title="자동 견적서 생성기", layout="wide")

# 1. 데이터 로드 함수
@st.cache_data(ttl=60) # 1분마다 엑셀 변경사항 체크
def load_data():
    if os.path.exists("products.xlsx"):
        return pd.read_excel("products.xlsx")
    else:
        st.error("파일을 찾을 수 없습니다. 'products.xlsx' 파일을 확인해주세요.")
        return pd.DataFrame()

df = load_data()

st.title("📑 스마트 견적서 관리 프로그램")

if not df.empty:
    # 2. 검색 및 조회 필터
    with st.sidebar:
        st.header("🔍 제품 검색")
        category = st.selectbox("분류 선택", ["전체"] + list(df['분류'].unique()))
        search_name = st.text_input("제품명(국문/영문) 검색")

    # 필터링
    filtered_df = df.copy()
    if category != "전체":
        filtered_df = filtered_df[filtered_df['분류'] == category]
    if search_name:
        filtered_df = filtered_df[filtered_df['품명(국문)'].str.contains(search_name) | filtered_df['품명(영문)'].str.contains(search_name)]

    st.subheader("📦 등록 제품 리스트")
    st.dataframe(filtered_df, use_container_width=True)

    # 3. 견적서 생성 섹션
    st.divider()
    st.subheader("✍️ 견적서 작성")
    
    selected_item = st.selectbox("견적을 작성할 제품을 선택하세요", filtered_df['품명(영문)'].tolist())
    
    if selected_item:
        item_data = df[df['품명(영문)'] == selected_item].iloc[0]
        
        col1, col2, col3 = st.columns([1, 2, 2])
        
        with col1:
            # 이미지 표시
            img_path = f"images/{item_data['이미지']}"
            if os.path.exists(img_path):
                st.image(img_path, caption=selected_item)
            else:
                st.warning("이미지 없음")

        with col2:
            st.write(f"**규격:** {item_data['규격(g)']}g")
            st.write(f"**수량/박스:** {item_data['수량/박스']} EA")
            st.write(f"**CBM:** {item_data['Weight CBM/CTN - CBM']}")
            st.write(f"**MOQ:** {item_data['MOQ']}")

        with col3:
            # 오퍼가 수정 영역
            new_offer_unit = st.number_input("오퍼가 FOB - 단가 (수정)", value=float(item_data['오퍼가 FOB -단가']))
            new_offer_ctn = st.number_input("오퍼가 FOB - C/T가격 (수정)", value=float(item_data['오퍼가 FOB-C/T가격']))

        # 4. 엑셀 다운로드 버튼
        if st.button("🚀 이 양식으로 견적서 다운로드 (Excel)"):
            output = BytesIO()
            workbook = xlsxwriter.Workbook(output)
            sheet = workbook.add_worksheet("Quotation")

            # 셀 서식
            header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9EAD3'})
            base_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})

            # 헤더 작성 (병합 포함)
            headers = [
                "Product Name(English)", "PICTURE", "Weight(EA)", "EA/CTN",
                "Weight, Cbm/ctn", "Weight, Cbm/ctn", "Weight, Cbm/ctn",
                "FOB KOREAN PORT", "FOB KOREAN PORT", "Storage", "Shelf Life", "MOQ"
            ]
            sub_headers = ["", "", "g", "", "net(kg)", "gross(kg)", "cbm", "EA", "CTN", "", "", ""]

            # 상위 헤더 병합 작성
            sheet.merge_range('A1:A2', headers[0], header_fmt)
            sheet.merge_range('B1:B2', headers[1], header_fmt)
            sheet.merge_range('C1:C2', headers[2], header_fmt)
            sheet.merge_range('D1:D2', headers[3], header_fmt)
            sheet.merge_range('E1:G1', headers[4], header_fmt) # Weight 그룹
            sheet.merge_range('H1:I1', headers[7], header_fmt) # FOB 그룹
            sheet.merge_range('J1:J2', headers[9], header_fmt)
            sheet.merge_range('K1:K2', headers[10], header_fmt)
            sheet.merge_range('L1:L2', headers[11], header_fmt)

            # 하위 헤더 작성
            for col, text in enumerate(sub_headers):
                if text: sheet.write(1, col, text, header_fmt)

            # 데이터 채우기 (3행)
            row = 2
            sheet.set_row(row, 80) # 행 높이 조절 (이미지용)
            sheet.write(row, 0, item_data['품명(영문)'], base_fmt)
            
            # 이미지 삽입
            if os.path.exists(img_path):
                sheet.insert_image(row, 1, img_path, {'x_scale': 0.1, 'y_scale': 0.1, 'x_offset': 5, 'y_offset': 5})
            
            sheet.write(row, 2, item_data['규격(g)'], base_fmt)
            sheet.write(row, 3, item_data['수량/박스'], base_fmt)
            sheet.write(row, 4, item_data['Weight CBM/CTN - net'], base_fmt)
            sheet.write(row, 5, item_data['Weight CBM/CTN - gross'], base_fmt)
            sheet.write(row, 6, item_data['Weight CBM/CTN - CBM'], base_fmt)
            sheet.write(row, 7, new_offer_unit, base_fmt)
            sheet.write(row, 8, new_offer_ctn, base_fmt)
            sheet.write(row, 9, item_data['storage'], base_fmt)
            sheet.write(row, 10, item_data['shelf life'], base_fmt)
            sheet.write(row, 11, item_data['MOQ'], base_fmt)

            workbook.close()
            
            st.download_button(
                label="📥 엑셀 파일 저장하기",
                data=output.getvalue(),
                file_name=f"Quotation_{selected_item}.xlsx",
                mime="application/vnd.ms-excel"
            )
else:
    st.info("데이터가 없습니다. products.xlsx 파일을 작성해주세요.")
