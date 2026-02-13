import pandas as pd
from .utils import clean_string

def process_delivery_data(file_storage):
    """
    Trích xuất thông tin phương tiện từ file BKPX.
    Trả về một từ điển mapping: { 'Số hóa đơn': 'Biển số xe' }
    """
    delivery_map = {}
    if not file_storage:
        return delivery_map

    try:
        try:
            df_raw = pd.read_excel(file_storage, header=None)
        except:
            file_storage.seek(0)
            df_raw = pd.read_csv(file_storage, header=None)

        # 1. Tìm dòng tiêu đề (Chứa 'Số hóa đơn' hoặc 'Phương tiện')
        header_idx = -1
        for i, row in df_raw.iterrows():
            row_vals = [str(c).strip().upper() for c in row if c is not None]
            if 'PHƯƠNG TIỆN' in row_vals or 'BIỂN SỐ' in row_vals:
                header_idx = i
                break
        
        if header_idx == -1:
            return delivery_map

        # 2. Ánh xạ cột
        header_row = df_raw.iloc[header_idx]
        col_inv = -1
        col_truck = -1
        
        for idx, val in enumerate(header_row):
            clean_val = str(val).strip().upper()
            if 'HÓA ĐƠN' in clean_val or 'HĐ' in clean_val:
                col_inv = idx
            if 'PHƯƠNG TIỆN' in clean_val or 'BIỂN SỐ' in clean_val or 'XE' in clean_val:
                col_truck = idx

        # 3. Thu thập dữ liệu
        if col_inv != -1 and col_truck != -1:
            # Bắt đầu đọc từ sau dòng tiêu đề
            df_data = df_raw.iloc[header_idx + 1:]
            for _, row in df_data.iterrows():
                raw_inv = str(row.iloc[col_inv]).strip()
                # Chuẩn hóa số hóa đơn để khớp (thường lấy 5-7 số cuối)
                if raw_inv and raw_inv.lower() != 'nan':
                    # Lấy 5 số cuối của số hóa đơn để làm Key khớp nối
                    inv_key = raw_inv.split('.')[0][-5:].zfill(5)
                    truck_no = clean_string(row.iloc[col_truck])
                    delivery_map[inv_key] = truck_no

    except Exception as e:
        print(f"Lỗi delivery_handler: {e}")
        
    return delivery_map