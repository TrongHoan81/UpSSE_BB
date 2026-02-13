import pandas as pd
from .utils import clean_string, to_float

def process_bm19_data(file_storage):
    """
    Trích xuất dữ liệu kỹ thuật từ BM19 (Nhiệt độ, Tỷ trọng).
    Cột quy định: F (Hàng), G (Khách), H (Nhiệt độ), I (Tỷ trọng), M (Số lượng).
    """
    bm19_list = []
    if not file_storage: return bm19_list

    try:
        df = pd.read_excel(file_storage, header=None)
        
        # Tìm dòng tiêu đề
        header_idx = -1
        for i, row in df.iterrows():
            vals = [str(c).strip().upper() for c in row if c is not None]
            if 'MẶT HÀNG' in vals and 'KHÁCH HÀNG' in vals:
                header_idx = i
                break
        
        if header_idx == -1: return []

        # Nạp dữ liệu với các khóa chuẩn: temp, dens, so_luong
        for idx in range(header_idx + 1, len(df)):
            row = df.iloc[idx]
            mat_hang = clean_string(row.iloc[5])
            if not mat_hang: continue
            
            bm19_list.append({
                'mat_hang': mat_hang.upper(),
                'ten_kh': str(row.iloc[6]).strip(), 
                'temp': row.iloc[7],  # Nhiệt độ hiện tại
                'dens': row.iloc[8],  # Tỷ trọng
                'so_luong': round(to_float(row.iloc[12]), 3) # Làm tròn để so khớp số lượng
            })
            
    except Exception as e:
        print(f"Lỗi khi xử lý BM19: {e}")
        
    return bm19_list