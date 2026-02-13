import os
from flask import Flask, render_template, request, send_file, jsonify
from handlers.invoice_handler import process_invoice_data
from handlers.bm19_handler import process_bm19_data 
from handlers.merge_handler import merge_and_fill_template

app = Flask(__name__)
app.secret_key = "pvoil_namdinh_upsse_bb_final_2026"

# Đảm bảo thư mục Data tồn tại
if not os.path.exists('Data'):
    os.makedirs('Data')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        file_bkhd = request.files.get('file_bkhd')
        file_bm19 = request.files.get('file_bm19')
        
        # Nhận ngày xác nhận từ Modal giao diện (nếu có)
        confirmed_date = request.form.get('confirmed_date')
        
        if not file_bkhd:
            return jsonify({'status': 'error', 'message': 'Vui lòng chọn tệp BKHD!'})

        # 1. Trích xuất BKHD và kiểm tra tính nhập nhèm ngày tháng
        invoice_result = process_invoice_data(file_bkhd)
        
        # Nếu phát hiện ngày nhập nhèm (Ví dụ 02/11) và chưa có xác nhận từ người dùng -> Hiện Modal
        if invoice_result['is_ambiguous'] and not confirmed_date:
            return jsonify({
                'status': 'ambiguous',
                'opt1': invoice_result['opt1'],
                'opt2': invoice_result['opt2']
            })

        # 2. Chốt ngày hóa đơn (Ưu tiên ngày đã người dùng xác nhận qua Modal)
        final_date = confirmed_date if confirmed_date else invoice_result['auto_date']

        # 3. Trích xuất dữ liệu BM19 (Nhiệt độ, Tỷ trọng) nếu người dùng có tải lên
        bm19_data = process_bm19_data(file_bm19) if file_bm19 else []
        
        # Đường dẫn file mẫu Template
        template_path = os.path.join('Data', 'template_svdetail9.xlsx')
        
        # 4. Thực hiện ghép nối, tính toán tách dòng và tạo file kết quả
        output_buffer = merge_and_fill_template(
            invoice_data=invoice_result['data'],
            bm19_data=bm19_data, 
            template_path=template_path,
            manual_date=final_date
        )

        return send_file(
            output_buffer,
            as_attachment=True,
            download_name="Ket_Qua_UpSSE_BB.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)