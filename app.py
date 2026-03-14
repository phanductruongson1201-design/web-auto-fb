from flask import Flask, request, render_template_string
import requests
import re
import pandas as pd
import os
import time

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Giao diện HTML đã được thiết kế lại (CSS Mới)
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ Thống Auto Đăng Bài Pro</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #4f46e5;
            --primary-hover: #4338ca;
            --bg-color: #f1f5f9;
            --card-bg: #ffffff;
            --text-main: #1e293b;
            --text-muted: #64748b;
            --border-color: #cbd5e1;
            --success: #10b981;
            --error: #ef4444;
        }
        body { 
            font-family: 'Inter', sans-serif; 
            margin: 0; 
            padding: 40px 20px; 
            background-color: var(--bg-color); 
            color: var(--text-main);
            display: flex;
            justify-content: center;
        }
        .container { 
            background: var(--card-bg); 
            padding: 40px; 
            border-radius: 16px; 
            width: 100%;
            max-width: 650px; 
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05); 
        }
        .header {
            text-align: center;
            margin-bottom: 35px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 20px;
        }
        .header h2 {
            margin: 0;
            color: var(--primary);
            font-size: 26px;
            font-weight: 700;
        }
        .header p {
            margin: 8px 0 0;
            color: var(--text-muted);
            font-size: 15px;
        }
        .form-group {
            margin-bottom: 22px;
        }
        label {
            display: block;
            font-weight: 600;
            margin-bottom: 8px;
            font-size: 14px;
            color: #334155;
        }
        .optional { font-size: 12px; color: var(--text-muted); font-weight: 400; font-style: italic; }
        input[type="text"], input[type="number"], textarea, input[type="file"] { 
            width: 100%; 
            padding: 12px 16px; 
            border: 1px solid var(--border-color); 
            border-radius: 8px; 
            box-sizing: border-box; 
            font-family: 'Inter', sans-serif;
            font-size: 14px;
            background-color: #f8fafc;
            transition: all 0.2s ease;
        }
        input[type="file"] {
            padding: 9px 16px;
        }
        input:focus, textarea:focus {
            outline: none;
            border-color: var(--primary);
            background-color: #ffffff;
            box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.15);
        }
        button { 
            background-color: var(--primary); 
            color: white; 
            padding: 15px 20px; 
            border: none; 
            border-radius: 8px; 
            cursor: pointer; 
            font-size: 16px; 
            font-weight: 600;
            width: 100%; 
            margin-top: 10px;
            transition: background-color 0.2s ease, transform 0.1s ease;
            box-shadow: 0 4px 6px -1px rgba(79, 70, 229, 0.3);
        }
        button:hover { 
            background-color: var(--primary-hover); 
        }
        button:active {
            transform: translateY(2px);
            box-shadow: 0 0px 0px rgba(0,0,0,0);
        }
        .result-box { 
            margin-top: 35px; 
            padding: 25px; 
            border-radius: 12px; 
            background-color: #f8fafc; 
            border: 1px solid var(--border-color);
        }
        .result-box h3 { margin-top: 0; font-size: 18px; margin-bottom: 20px; color: #0f172a;}
        .result-list { list-style: none; padding: 0; margin: 0; }
        .result-item { 
            padding: 12px 16px; 
            margin-bottom: 10px; 
            border-radius: 8px; 
            font-size: 14px;
            border-left: 5px solid transparent;
            font-weight: 500;
        }
        .success { background-color: #ecfdf5; border-left-color: var(--success); color: #065f46; border-right: 1px solid #d1fae5; border-top: 1px solid #d1fae5; border-bottom: 1px solid #d1fae5;}
        .error { background-color: #fef2f2; border-left-color: var(--error); color: #991b1b; border-right: 1px solid #fee2e2; border-top: 1px solid #fee2e2; border-bottom: 1px solid #fee2e2;}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>Hệ Thống Auto Đăng Bài Pro</h2>
            <p>Trình quản lý đăng bài hàng loạt qua API Chính thức</p>
        </div>
        <form method="POST" enctype="multipart/form-data">
            <div class="form-group">
                <label>Mã Access Token (Bắt buộc):</label>
                <input type="text" name="access_token" required placeholder="Nhập chuỗi token từ Meta Developers...">
            </div>
            
            <div class="form-group">
                <label>Tệp danh sách Excel (.xlsx):</label>
                <input type="file" name="excel_file" accept=".xlsx, .xls" required>
            </div>
            
            <div class="form-group">
                <label>Hình ảnh đính kèm <span class="optional">(Không bắt buộc)</span>:</label>
                <input type="file" name="image_file" accept="image/*">
            </div>
            
            <div class="form-group">
                <label>Nội dung bài viết:</label>
                <textarea name="message" required rows="6" placeholder="Nhập nội dung bạn muốn truyền tải đến các hội nhóm..."></textarea>
            </div>
            
            <div class="form-group">
                <label>Khoảng nghỉ chống Spam (Giây):</label>
                <input type="number" name="delay" value="5" min="1" required>
            </div>
            
            <button type="submit">Khởi Chạy Tiến Trình</button>
        </form>

        {% if results %}
            <div class="result-box">
                <h3>Bảng Báo Cáo Kết Quả</h3>
                <ul class="result-list">
                {% for res in results %}
                    <li class="result-item {{ 'success' if '✅' in res.status else 'error' }}">
                        <strong>ID Nhóm: {{ res.group }}</strong> <br> 
                        <span style="font-weight: 400; font-size: 13px; margin-top: 4px; display: inline-block;">Trạng thái: {{ res.status }}</span>
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
        excel_file = request.files.get('excel_file')
        image_file = request.files.get('image_file')

        image_bytes = None
        image_name = None
        if image_file and image_file.filename != '':
            image_bytes = image_file.read()
            image_name = image_file.filename

        if excel_file and excel_file.filename != '':
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
            excel_file.save(filepath)

            try:
                df = pd.read_excel(filepath)
                group_list = df.iloc[:, 0].dropna().tolist()

                for item in group_list:
                    group_id = extract_group_id(item)
                    if not group_id:
                        results.append({'group': item, 'status': '❌ Không trích xuất được ID từ Link'})
                        continue

                    if image_bytes:
                        url = f"https://graph.facebook.com/v19.0/{group_id}/photos"
                        payload = {'message': message, 'access_token': access_token}
                        files = {'source': (image_name, image_bytes, 'image/jpeg')}
                        response = requests.post(url, data=payload, files=files)
                    else:
                        url = f"https://graph.facebook.com/v19.0/{group_id}/feed"
                        payload = {'message': message, 'access_token': access_token}
                        response = requests.post(url, data=payload)
                    
                    try:
                        data = response.json()
                        if 'id' in data or 'post_id' in data:
                            results.append({'group': group_id, 'status': f"✅ Đăng thành công (Mã bài: {data.get('post_id', data.get('id'))})"})
                        else:
                            error_msg = data.get('error', {}).get('message', 'Lỗi không xác định')
                            results.append({'group': group_id, 'status': f"❌ Bị từ chối: {error_msg}"})
                    except Exception as e:
                        results.append({'group': group_id, 'status': f"❌ Lỗi kết nối máy chủ FB: {str(e)}"})
                    
                    time.sleep(delay)

            except Exception as e:
                results.append({'group': 'Hệ thống', 'status': f"❌ Lỗi xử lý file Excel: {str(e)}"})
            finally:
                if os.path.exists(filepath):
                    os.remove(filepath)
                
    return render_template_string(HTML_TEMPLATE, results=results)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
