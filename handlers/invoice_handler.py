import pandas as pd
from datetime import datetime, timedelta
from .utils import clean_string, to_float, normalize_mst

def check_date_ambiguity(excel_serial):
    """
    Phát hiện các trường hợp ngày tháng nhập nhèm (D và M đều <= 12).
    Trả về: (is_ambiguous, option1, option2, auto_date_str)
    """
    if not excel_serial or str(excel_serial).lower() == 'nan':
        return False, None, None, ""
    
    try:
        # Chuyển đổi Serial Date của Excel sang Python Datetime
        base_date = datetime(1899, 12, 30)
        dt = base_date + timedelta(days=int(float(excel_serial)))
        
        d, m, y = dt.day, dt.month, dt.year
        
        # Logic: Nếu có giá trị > 12 -> Tự hiểu giá trị đó là Ngày
        if d > 12 or m > 12:
            if m > 12: # Nếu Excel đang hiểu Tháng là Ngày (>12)
                return False, None, None, f"{m:02d}/{d:02d}/{y}"
            return False, None, None, f"{d:02d}/{m:02d}/{y}"
        
        # Nếu cả hai đều <= 12 và khác nhau -> Cần người dùng xác nhận
        if d <= 12 and m <= 12 and d != m:
            return True, {"d": d, "m": m, "full": f"{d:02d}/{m:02d}/{y}"}, \
                         {"d": m, "m": d, "full": f"{m:02d}/{d:02d}/{y}"}, ""
        
        # Trường hợp d == m (Ví dụ 05/05)
        return False, None, None, f"{d:02d}/{m:02d}/{y}"
    except:
        return False, None, None, str(excel_serial)

def process_invoice_data(file_storage):
    """Trích xuất dữ liệu BKHD và xác định tính rõ ràng của ngày tháng."""
    try:
        df_raw = pd.read_excel(file_storage, header=None)
        
        # Quét tìm dòng tiêu đề chứa 'STT'
        header_idx = -1
        for i, row in df_raw.iterrows():
            row_vals = [str(c).strip().upper() for c in row[:15] if c is not None]
            if 'STT' in row_vals:
                header_idx = i
                break
        
        if header_idx == -1:
            raise ValueError("Không tìm thấy dòng tiêu đề 'STT' trong file BKHD.")

        header_row = df_raw.iloc[header_idx]
        col_map = {}
        targets = {
            'TEN_KH': 'TÊN KHÁCH HÀNG', 'MST': 'MST KHÁCH HÀNG', 'MAT_HANG': 'MẶT HÀNG',
            'KHO_XUAT': 'KHO XUẤT HÀNG', 'SO_LUONG': 'SỐ LƯỢNG', 'DON_GIA': 'ĐƠN GIÁ',
            'DVT': 'ĐƠN VỊ TÍNH', 'THANH_TIEN': 'THÀNH TIỀN', 'VAT_RATE': 'VAT',
            'TIEN_THUE_BKHD': 'TIỀN THUẾ', 'MAU_SO': 'MẪU SỐ', 'KY_HIEU': 'KÝ HIỆU',
            'SO_HD': 'SỐ HÓA ĐƠN', 'NGAY_HD': 'NGÀY HÓA ĐƠN'
        }

        for idx, val in enumerate(header_row):
            cv = str(val).strip().upper()
            for key, search_name in targets.items():
                if key not in col_map and search_name in cv:
                    col_map[key] = idx

        # Phân tích ngày từ dòng dữ liệu đầu tiên
        first_row_data = df_raw.iloc[header_idx + 2]
        is_amb, o1, o2, auto_d = check_date_ambiguity(first_row_data.iloc[col_map['NGAY_HD']])

        # Thu thập dữ liệu các dòng hóa đơn
        data_list = []
        for _, row in df_raw.iloc[header_idx + 2:].iterrows():
            stt = str(row.iloc[0]).strip().lower()
            if not stt or stt == 'nan' or 'cộng' in stt: continue
            
            data_list.append({
                'ten_kh': clean_string(row.iloc[col_map['TEN_KH']]),
                'mst_key': normalize_mst(row.iloc[col_map['MST']]),
                'so_hd': clean_string(row.iloc[col_map['SO_HD']]),
                'ky_hieu': clean_string(row.iloc[col_map['KY_HIEU']]),
                'mau_so': clean_string(row.iloc[col_map['MAU_SO']]),
                'dvt': clean_string(row.iloc[col_map['DVT']]),
                'so_luong': to_float(row.iloc[col_map['SO_LUONG']]),
                'don_gia_bkhd': to_float(row.iloc[col_map['DON_GIA']]),
                'vat_raw': row.iloc[col_map['VAT_RATE']],
                'tien_thue_total_bkhd': to_float(row.iloc[col_map['TIEN_THUE_BKHD']]),
                'mat_hang': clean_string(row.iloc[col_map['MAT_HANG']]),
                'kho_xuat_bkhd': clean_string(row.iloc[col_map.get('KHO_XUAT', 0)]),
                'thanh_tien_total_bkhd': to_float(row.iloc[col_map['THANH_TIEN']])
            })

        return {
            'data': data_list, 
            'is_ambiguous': is_amb, 
            'opt1': o1, 
            'opt2': o2, 
            'auto_date': auto_d
        }
    except Exception as e:
        raise ValueError(f"Lỗi khi trích xuất BKHD: {str(e)}")