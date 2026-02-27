import re
import math

def clean_string(s):
    """Xóa khoảng trắng thừa và ký tự rác."""
    if s is None or str(s).lower() == 'nan': return ""
    cleaned = str(s).strip()
    if cleaned.startswith("'"): cleaned = cleaned[1:]
    return re.sub(r'\s+', ' ', cleaned)

def normalize_name_advanced(name):
    """
    Chuẩn hóa tên khách hàng chuyên sâu: Giải mã viết tắt (TM, DV, XN, BX, ĐT, XD...)
    để đồng bộ tên khách giữa BKHD và BM19.
    """
    if not name: return ""
    s = str(name).upper()
    
    replacements = {
        r'\bTM\b': 'THUONG MAI', r'\bTHƯƠNG MẠI\b': 'THUONG MAI',
        r'\bDV\b': 'DICH VU', r'\bDỊCH VỤ\b': 'DICH VU',
        r'\bTNHH\b': 'TRACH NHIEM HUU HAN',
        r'\bCP\b': 'CO PHAN', r'\bMTV\b': 'MOT THANH VIEN',
        r'\bVT\b': 'VAN TAI', r'\bVẬN TẢI\b': 'VAN TAI',
        r'\bXN\b': 'XI NGIEP', r'\bXÍ NGHIỆP\b': 'XI NGIEP',
        r'\bBX\b': 'BEN XE', r'\bBẾN XE\b': 'BEN XE',
        r'\bĐT\b': 'DAU TU', r'\bĐẦU TƯ\b': 'DAU TU',
        r'\bXD\b': 'XANG DAU', r'\bXĂNG DẦU\b': 'XANG DAU',
        r'\bXÂY DỰNG\b': 'XAY DUNG',
        r'\bDOANH NGHIỆP\b': 'DN',
        r'\bDN\b': 'DOANH NGHIEP'
    }
    for pattern, replacement in replacements.items():
        s = re.sub(pattern, replacement, s)
        
    s = re.sub(r'[^A-Z0-9 ]', '', s)
    return " ".join(s.split())

def calculate_similarity(name1, name2):
    """Tính điểm tương đồng Token-based Jaccard."""
    if not name1 or not name2: return 0.0
    s1 = set(normalize_name_advanced(name1).split())
    s2 = set(normalize_name_advanced(name2).split())
    if not s1 or not s2: return 0.0
    return len(s1.intersection(s2)) / len(s1.union(s2))

def to_float(value):
    """Ép kiểu số thực an toàn."""
    if value is None: return 0.0
    try:
        s_val = str(value).replace(',', '').strip()
        return float(s_val) if s_val and s_val.lower() != 'nan' else 0.0
    except: return 0.0

def to_tax_rate_float(raw_vat):
    """Quy đổi mã thuế suất."""
    if not raw_vat or str(raw_vat).lower() == 'nan': return 0.0
    try:
        s_val = str(raw_vat).replace('%', '').strip()
        f_val = float(s_val)
        return f_val / 100 if f_val >= 1 else f_val
    except: return 0.0

def normalize_mst(mst):
    """Chuẩn hóa mã số thuế."""
    if mst is None: return ""
    return re.sub(r'[^A-Z0-9]', '', str(mst).upper())

def format_tax_code(raw_vat):
    """Mã thuế 2 chữ số (08, 10)."""
    if not raw_vat or str(raw_vat).lower() == 'nan': return ""
    try:
        s_val = str(raw_vat).replace('%', '').strip()
        f_val = float(s_val)
        if 0 < f_val < 1: f_val *= 100
        return f"{int(round(f_val)):02d}"
    except: return str(raw_vat)