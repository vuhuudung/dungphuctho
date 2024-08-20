# Cài đặt thư viện
import streamlit as st
import pandas as pd
import numpy as np
import unidecode
import os
import requests
import io
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder
import webbrowser
from pathlib import Path

def render():
    # Thiết lập wide mode cho Streamlit
    st.set_page_config(layout="wide")

    st.header('TRA CỨU VĂN BẢN 2024')
    st.text('tra cứu văn bản')
    st.text('@@@@@@@@@@@@@@@@')

####################################
    ten_file_excel_0= r'data/vanban.xlsx'
    ten_sheet_0='IncomingDocuments'  # Thay thế bằng tên của sheet bạn muốn thêm dữ liệu
    df=pd.read_excel(ten_file_excel_0, sheet_name=ten_sheet_0,  header=0, dtype=str) #lấy dữ liệu, sheet = ho so, không lấy header

    df=pd.DataFrame(df)
    df=df.astype(str)
    
    # Hàm loại bỏ dấu và chuyển về chữ thường
    def chu_thuong_khong_dau(text):
        return unidecode.unidecode(text).lower()

    # Hàm chuyển đổi danh sách
    def chuyen_doi_chu_thuong(lst):
        return [chu_thuong_khong_dau(item) for item in lst]
    
    # Nhập dữ liệu tìm kiếm
    trich_yeu = 'Về việc' # str.contains để kiểm tra hoặc có có chuỗi này không
    loai_bo_trich_yeu = '@@@'
    so_ky_hieu = 'qđ' #.apply để kiểm tra có đồng thời các chuỗi này không
    loai_bo_so_ky_hieu = '@@@'
    ngay_van_ban = '2024' # str.contains để kiểm tra hoặc có có chuỗi này không

    # Add 2 columns for date input
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        trich_yeu = st.text_input('Trích yếu', trich_yeu)

    with col2:
        loai_bo_trich_yeu = st.text_input('Loại bỏ nội dung', loai_bo_trich_yeu)

    with col3:
        so_ky_hieu = st.text_input('Số ký hiệu ', so_ky_hieu)

    with col4:
        loai_bo_so_ky_hieu = st.text_input('Loại bỏ số ký hiệu', loai_bo_so_ky_hieu)

    with col5:
        ngay_van_ban = st.text_input('Ngày văn bản', ngay_van_ban)

    # chuyển đổi kiểu chữ   
    trich_yeu = chuyen_doi_chu_thuong([item.strip() for item in trich_yeu.split(',')])
    loai_bo_trich_yeu = chuyen_doi_chu_thuong([item.strip() for item in loai_bo_trich_yeu.split(',')])
    so_ky_hieu = chuyen_doi_chu_thuong([item.strip() for item in so_ky_hieu.split(',')])
    loai_bo_so_ky_hieu = chuyen_doi_chu_thuong([item.strip() for item in loai_bo_so_ky_hieu.split(',')])
    ngay_van_ban = chuyen_doi_chu_thuong([item.strip() for item in ngay_van_ban.split(',')])

    # Tìm kiếm và lọc các dòng chứa ít nhất một trong các chuỗi cần lọc
    df_loc = df[df['Trích yếu'].apply(chu_thuong_khong_dau).str.contains('|'.join(trich_yeu),case=False) 
                & ~df['Trích yếu'].apply(chu_thuong_khong_dau).str.contains('|'.join(loai_bo_trich_yeu), na=False)
                & df['Số ký hiệu'].apply(chu_thuong_khong_dau).apply(lambda x: all(chuoi in x for chuoi in so_ky_hieu))
                & ~df['Số ký hiệu'].apply(chu_thuong_khong_dau).str.contains('|'.join(loai_bo_so_ky_hieu), na=False)
                & df['Ngày văn bản'].apply(chu_thuong_khong_dau).str.contains('|'.join(ngay_van_ban),case=False) 
    ]
    
    # Hiển thị kết quả
    df_hienthi = df_loc[['STT', 'Trích yếu', 'Số ký hiệu', 'Ngày văn bản', 'Đơn vị ban hành', 'Ghi chú']]
    
    # Sử dụng st_aggrid để tùy chỉnh chiều rộng của từng cột
    gb = GridOptionsBuilder.from_dataframe(df_hienthi)

    # Cấu hình chiều rộng của từng cột
    gb.configure_column("STT", width=10, sortable=True)
    gb.configure_column("Trích yếu", width=100, sortable=True)
    gb.configure_column("Số ký hiệu", width=10, sortable=True)
    gb.configure_column("Ngày văn bản", width=10, sortable=True)
    gb.configure_column("Đơn vị ban hành", width=10, sortable=True)
    gb.configure_column("Ghi chú", width=10, sortable=True)
    gb.configure_column("Path", width=10, sortable=True)

    # Cấu hình phân trang
    gb.configure_pagination(paginationAutoPageSize=True)  # Chia thành nhiều trang tự động
    gb.configure_default_column(wrapText=True, autoHeight=True)  # Tự động giãn ô khi chuỗi dài
    gridOptions = gb.build()

    # Hiển thị DataFrame với cấu hình tùy chỉnh
    st.write("Bảng dữ liệu:")
    AgGrid(df_hienthi, gridOptions=gridOptions, height=400, fit_columns_on_grid_load=True)
        # Tạo nút tải xuống Excel
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_hienthi.to_excel(writer, index=False, sheet_name='Dữ liệu')
    excel_data = excel_buffer.getvalue()

    st.download_button(
        label="Tải xuống Excel",
        data=excel_data,
        file_name='du_lieu.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ) 

####################################
    path_file= r'data/vanban1.xlsx'
    path_file_sheet='IncomingDocuments'  # Thay thế bằng tên của sheet bạn muốn thêm dữ liệu
    df_path=pd.read_excel(path_file, sheet_name=path_file_sheet,  header=0, dtype=str) #lấy dữ liệu, sheet = ho so, không lấy header

    df_path=pd.DataFrame(df_path)
    df_path=df_path.astype(str)

    so_ky_hieu_2 = '000, qldt, 2024' #.apply để kiểm tra có đồng thời các chuỗi này không
    so_ky_hieu_2 = st.text_input('Số ký hiệu tên thư mục ', so_ky_hieu_2)
    so_ky_hieu_2 = chuyen_doi_chu_thuong([item.strip() for item in so_ky_hieu_2.split(',')])
    # Tìm kiếm và lọc các dòng chứa ít nhất một trong các chuỗi cần lọc
    df_path_loc = df_path[df_path['Folder Name'].apply(chu_thuong_khong_dau).apply(lambda x: all(chuoi in x for chuoi in so_ky_hieu_2))]
    
    # Sử dụng st_aggrid để tùy chỉnh chiều rộng của từng cột
    gb1 = GridOptionsBuilder.from_dataframe(df_path_loc)

    # Cấu hình chiều rộng của từng cột
    gb1.configure_column("Folder Name", width=10, sortable=True)
    gb1.configure_column("Full Path", width=20, sortable=True)
    # Cấu hình phân trang
    gb1.configure_pagination(paginationAutoPageSize=True)  # Chia thành nhiều trang tự động
    gb1.configure_default_column(wrapText=True, autoHeight=True)  # Tự động giãn ô khi chuỗi dài
    gridOptions1 = gb1.build()
    # Hiển thị DataFrame với cấu hình tùy chỉnh
    # st.write("Bảng dữ liệu:")
    # AgGrid(df_path_loc, gridOptions=gridOptions1, height=200, fit_columns_on_grid_load=True)

########################################
    # Hàm mở đường dẫn local
    def open_local_path(path):
        if os.path.exists(path):
            os.startfile(path)  # Đối với Windows
            # webbrowser.open(path)  # Đối với các hệ điều hành khác
        else:
            st.error(f"Đường dẫn không tồn tại: {path}")

    # # Kiểm tra và lưu trạng thái danh sách button trong session_state
    # if 'show_buttons' not in st.session_state:
    #     st.session_state.show_buttons = False
    
    # # Button để mở hoặc đóng danh sách các button
    # if st.button("Hiển thị/Ẩn danh sách đường dẫn"):
    #     st.session_state.show_buttons = not st.session_state.show_buttons


    # # # Hiển thị danh sách các button nếu show_buttons là True
    # if st.session_state.show_buttons:
    #     st.write("Chọn đường dẫn để mở:")
    #     for index, row in df_path_loc.iterrows():
    #         col1, col2, col3 = st.columns([2, 4, 1])
    #         with col1:
    #             st.write(row['Folder Name'])
    #         with col2:
    #             st.write(row['Full Path'])
    #         with col3:
    #             if st.button('Mở File', key=index):
    #                 open_local_path(row['Full Path'])
    st.write("Chọn đường dẫn để mở:")
    for index, row in df_path_loc.iterrows():
        col1, col2, col3 = st.columns([2, 4, 1])
        with col1:
            st.write(row['Folder Name'])
        with col2:
            st.write(row['Full Path'])
        with col3:
            if st.button('Mở File', key=index):
                open_local_path(row['Full Path'])
        
render()
