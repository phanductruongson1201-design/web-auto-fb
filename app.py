from flask import Flask, request, render_template_string
import requests
import re
import pandas as pd
import os
import time

app = Flask(__name__)
# Tạo thư mục tạm để lưu file upload
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ thống Đăng bài Group API Hàng Loạt</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; background-color: #f0f2f5; }
        .container { background: white; padding: 20px; border-radius: 8px; max-width: 600px; margin: auto; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        input, textarea { width: 100%; padding: 10px; margin-top: 5px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        button { background-color: #0866ff; color: white; padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; width: 100%; }
        button:hover { background-color: #0054d1; }
        .result-box { margin-top: 20px; padding: 15px; border-radius: 4px; background-color: #e7f3ff; }
        .success { color: green; }
        .error { color: red; }
    </style>
</head>
<body>
    <div class="container">
        <h2>Đăng bài Facebook Group Hàng Loạt (Excel)</h2>
        <form method="POST" enctype="multipart/form-data">
            <label>Access Token (Bắt buộc):</label>
            <input type="text" name="access_token" required placeholder="Nhập mã token hợp lệ...">
            
            <label>Tải lên file Excel (.xlsx) chứa Link/ID Nhóm:</label>
            <input type="file" name="excel_file" accept=".xlsx, .xls" required>
            
            <label>Nội dung bài viết:</label>
            <textarea name="message" required rows="5" placeholder="Nhập nội dung bạn muốn đăng..."></textarea>
            
            <label>Thời gian nghỉ giữa các bài (Giây - Tránh bị khóa API):</label>
            <input type="number" name="delay" value="5" min="1" required>
            
            <button type="submit">Bắt đầu Đăng Bài</button>
        </form>

        {% if results %}
            <div class="result-box">
                <h3>Kết quả chi tiết:</h3>
                <ul>
                {% for res in results %}
                    <li class="{{ 'success' if '✅' in res.status else 'error' }}">
                        <strong>{{ res.group }}:</strong> {{ res.status }}
                    </li>
                {% endfor %}
                </ul>
            </div>
        {% endif %}
    </div>
</body>
</html>
"""

def extract_group_id(group_input):
    group_str = str(group_input).strip()
    if group_str.isdigit():
        return group_str
    match = re.search(r'groups/(\d+)', group_str)
    if match:
        return match.group(1)
    return None

@app.route('/', methods=['GET', 'POST'])
def index():
    results = []
    if request.method == 'POST':
        access_token = request.form['access_token']
        message = request.form['message']
        delay = int(request.form.get('delay', 5))
        file = request.files['excel_file']

        if file and file.filename != '':
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)

            try:
                # Đọc file Excel, lấy dữ liệu từ cột đầu tiên
                df = pd.read_excel(filepath)
                # Giả định cột đầu tiên chứa link/ID
                group_list = df.iloc[:, 0].dropna().tolist()

                for item in group_list:
                    group_id = extract_group_id(item)
                    if not group_id:
                        results.append({'group': item, 'status': '❌ Không trích xuất được ID'})
                        continue

                    # Gọi Graph API
                    url = f"https://graph.facebook.com/v19.0/{group_id}/feed"
                    payload = {
                        'message': message,
                        'access_token': access_token
                    }
                    
                    try:
                        response = requests.post(url, data=payload)
                        data = response.json()
                        
                        if 'id' in data:
                            results.append({'group': group_id, 'status': f"✅ Thành công (ID: {data['id']})"})
                        else:
                            error_msg = data.get('error', {}).get('message', 'Lỗi không xác định')
                            results.append({'group': group_id, 'status': f"❌ Bị từ chối: {error_msg}"})
                    except Exception as e:
                        results.append({'group': group_id, 'status': f"❌ Lỗi kết nối: {str(e)}"})
                    
                    # Tạm dừng để tránh bị Facebook chặn do gọi API quá nhanh (Rate Limiting)
                    time.sleep(delay)

            except Exception as e:
                results.append({'group': 'Hệ thống', 'status': f"❌ Lỗi đọc file Excel: {str(e)}"})
            finally:
                # Xóa file tạm sau khi xử lý xong
                if os.path.exists(filepath):
                    os.remove(filepath)
                
    return render_template_string(HTML_TEMPLATE, results=results)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
