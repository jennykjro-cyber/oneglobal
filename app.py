import streamlit as st
import pandas as pd
import os
from io import BytesIO
import xlsxwriter

# 페이지 설정
st.set_page_config(page_title="OneGlobal 스마트 견적기", layout="wide")

# 1. 데이터 로드 함수
@st.cache_data(ttl=5)
def load_data():
    if os.path.exists("products.xlsx"):
        df = pd.read_excel("products.xlsx")
        df.columns = df.columns.str.strip() # 컬럼명 공백 제거
        return df
    else:
        return pd.DataFrame()

df_origin = load_data()

st.title("📦 OneGlobal 품목 관리 시스템")

if not df_origin.empty:
    # --- 2. 검색 및 필터 영역 ---
    st.subheader("🔍 1. 제품 검색 및 가격 수정")
    st.info("💡 표에서 직접 가격을 수정하고 왼쪽 체크박스를 선택하세요.")
    
    c1, c2 = st.columns(2)
    with c1:
        categories = ["전체"] + sorted(df_origin['분류'].unique().tolist())
        sel_cat = st.selectbox("분류별 보기", categories)
    with c2:
        search_txt = st.text_input("제품명(국문) 또는 Description 검색")

    # 필터링 로직
    filtered_df = df_origin.copy()
    if sel_cat != "전체":
        filtered_df = filtered_df[filtered_df['분류'] == sel_cat]
    if search_txt:
        filtered_df = filtered_df[
            filtered_df['품명(국문)'].str.contains(search_txt, na=False) | 
            filtered_df['Description of Goods'].str.contains(search_txt, na=False)
        ]

    # 선택 및 수정용 컬럼 추가
    if "선택" not in filtered_df.columns:
        filtered_df.insert(0, "선택", False)
    
    # --- 3. 데이터 에디터 (여기서 수정 및 선택 수행) ---
    edited_df = st.data_editor(
        filtered_df,
        hide_index=True,
        column_config={
            "선택": st.column_config.CheckboxColumn("선택", default=False),
            "오퍼가 FOB -단가": st.column_config.NumberColumn("단가 ($)", format="$ %.2f"),
            "오퍼가 FOB-C/T가격": st.column_config.NumberColumn("C/T가 ($)", format="$ %.2f"),
            "이미지": None  # 이미지 경로는 굳이 편집할 필요 없으므로 숨김
        },
        use_container_width=True,
        key="main_editor"
    )

    # 선택된 행만 추출
    selected_items = edited_df[edited_df["선택"] == True]

    # --- 4. 견적서 생성 영역 ---
    st.divider()
    st.subheader("📝 2. 선택된 견적 항목 확인")

    if not selected_items.empty:
        st.write(f"현재 **{len(selected_items)}**개의 품목이 선택되었습니다.")
        st.dataframe(selected_items[['품명(국문)', 'Description of Goods', '오퍼가 FOB -단가', '오퍼가 FOB-C/T가격']], hide_index=True)

        if st.button("📊 수정된 내용으로 견적서(Excel) 다운로드"):
            output = BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            sheet = workbook.add_worksheet("Quotation")

            # 스타일 설정
            head_style = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9EAD3'})
            data_style = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
            money_style = workbook.add_format({'num_format': '$#,##0.00', 'align': 'center', 'valign': 'vcenter', 'border': 1})

            # 헤더 정의 (계층 구조 반영)
            # (상위분류, 하위분류, 컬럼키)
            header_map = [
                (None, 'Description of Goods', 'Description of Goods'),
                (None, 'PICTURE', '이미지'),
                (None, 'Weight(EA)', '규격(g)'),
                (None, 'EA/CTN', '수량/박스'),
                ('Weight, Cbm/ctn', 'net(kg)', 'Weight CBM/CTN - net'),
                ('Weight, Cbm/ctn', 'gross(kg)', 'Weight CBM/CTN - gross'),
                ('Weight, Cbm/ctn', 'cbm', 'Weight CBM/CTN - CBM'),
                ('FOB KOREAN PORT', 'EA ($)', '오퍼가 FOB -단가'),
                ('FOB KOREAN PORT', 'CTN ($)', '오퍼가 FOB-C/T가격'),
                (None, 'Storage', 'storage'),
                (None, 'Shelf Life', 'shelf life'),
                (None, 'MOQ', 'MOQ')
            ]

            # 헤더 작성 (2개 행 사용)
            for col, (parent, child, _) in enumerate(header_map):
                if parent:
                    # 상위 분류가 있는 경우 병합 시도 (이전 열과 같으면 스킵 로직은 단순화 위해 수동 지정 가능)
                    # 여기서는 직관적으로 4-6열, 7-8열 병합
                    if col == 4: sheet.merge_range(0, 4, 0, 6, parent, head_style)
                    if col == 7: sheet.merge_range(0, 7, 0, 8, parent, head_style)
                    sheet.write(1, col, child, head_style)
                else:
                    sheet.merge_range(0, col, 1, col, child, head_style)

            # 데이터 입력
            for row_idx, (_, item) in enumerate(selected_items.iterrows()):
                curr_row = row_idx + 2 # 헤더가 2줄이므로 2부터 시작
                sheet.set_row(curr_row, 80)
                
                for col_idx, (_, _, key) in enumerate(header_map):
                    val = item[key]
                    
                    if key == '이미지':
                        img_path = f"images/{val}"
                        if os.path.exists(img_path):
                            sheet.insert_image(curr_row, col_idx, img_path, {'x_scale': 0.1, 'y_scale': 0.1, 'x_offset': 5, 'y_offset': 5})
                        else:
                            sheet.write(curr_row, col_idx, "N/A", data_style)
                    elif '오퍼가' in key:
                        sheet.write(curr_row, col_idx, val, money_style)
                    elif key == '규격(g)':
                        sheet.write(curr_row, col_idx, f"{val}g", data_style)
                    else:
                        sheet.write(curr_row, col_idx, val, data_style)

            sheet.set_column('A:A', 30)
            sheet.set_column('B:L', 15)
            workbook.close()
            
            st.download_button(
                label="💾 엑셀 파일 받기",
                data=output.getvalue(),
                file_name="OneGlobal_Quotation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("검색 결과 표에서 견적서에 포함할 제품의 '선택' 칸을 체크해주세요. 가격 수정도 가능합니다.")
else:
    st.error("엑셀 파일(products.xlsx)이 없거나 데이터가 비어있습니다.")
