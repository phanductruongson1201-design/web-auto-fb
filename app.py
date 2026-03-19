from flask import Flask, request, render_template_string, redirect, url_for, session
import requests
import re
import os
import time
import openpyxl

app = Flask(__name__)
app.secret_key = os.urandom(24) # Khóa bảo mật để lưu phiên đăng nhập
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --- CẤU HÌNH FACEBOOK APP ---
APP_ID = '385078767129314' # ID ứng dụng của bạn
APP_SECRET = 'f9d18c2c52c07ad7fede00f56243cfc6' # Bạn hãy dán App Secret vào giữa 2 dấu nháy đơn này
REDIRECT_URI = 'https://he-thong-dang-bai.onrender.com/callback'
# -----------------------------

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ Thống Auto Đăng Bài Pro</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root { --primary: #4f46e5; --primary-hover: #4338ca; --bg-color: #f1f5f9; --card-bg: #ffffff; --text-main: #1e293b; --text-muted: #64748b; --border-color: #cbd5e1; --success: #10b981; --error: #ef4444; }
        body { font-family: 'Inter', sans-serif; margin: 0; padding: 40px 20px; background-color: var(--bg-color); color: var(--text-main); display: flex; justify-content: center; }
        .container { background: var(--card-bg); padding: 40px; border-radius: 16px; width: 100%; max-width: 650px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }
        .header { text-align: center; margin-bottom: 35px; border-bottom: 1px solid var(--border-color); padding-bottom: 20px; }
        .header h2 { margin: 0; color: var(--primary); font-size: 26px; font-weight: 700; }
        .header p { margin: 8px 0 0; color: var(--text-muted); font-size: 15px; }
        .form-group { margin-bottom: 22px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; font-size: 14px; color: #334155; }
        .optional { font-size: 12px; color: var(--text-muted); font-weight: 400; font-style: italic; }
        input[type="text"], input[type="number"], textarea, input[type="file"] { width: 100%; padding: 12px 16px; border: 1px solid var(--border-color); border-radius: 8px; box-sizing: border-box; font-family: 'Inter', sans-serif; font-size: 14px; background-color: #f8fafc; transition: all 0.2s ease; }
        input[type="file"] { padding: 9px 16px; }
        input:focus, textarea:focus { outline: none; border-color: var(--primary); background-color: #ffffff; box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.15); }
        button { background-color: var(--primary); color: white; padding: 15px 20px; border: none; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: 600; width: 100%; margin-top: 10px; transition: all 0.2s ease; }
        button:hover { background-color: var(--primary-hover); }
        .result-box { margin-top: 35px; padding: 25px; border-radius: 12px; background-color: #f8fafc; border: 1px solid var(--border-color); }
        .result-box h3 { margin-top: 0; font-size: 18px; margin-bottom: 20px; color: #0f172a;}
        .result-list { list-style: none; padding: 0; margin: 0; }
        .result-item { padding: 12px 16px; margin-bottom: 10px; border-radius: 8px; font-size: 14px; border-left: 5px solid transparent; font-weight: 500; word-break: break-all; }
        .success { background-color: #ecfdf5; border-left-color: var(--success); color: #065f46; border: 1px solid #d1fae5; border-left: 5px solid var(--success); }
        .error { background-color: #fef2f2; border-left-color: var(--error); color: #991b1b; border: 1px solid #fee2e2; border-left: 5px solid var(--error); }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>Hệ Thống Auto Đăng Bài Pro</h2>
            <p>Trình phân luồng Hội Nhóm & Fanpage Tự Động</p>
        </div>
        
        {% if not is_logged_in %}
            <div style="text-align: center; margin-top: 40px; margin-bottom: 40px;">
                <p style="margin-bottom: 20px; color: var(--text-muted); font-size: 15px;">Vui lòng kết nối với Facebook để hệ thống lấy quyền đăng bài tự động.</p>
                <a href="{{ login_url }}" style="background-color: #1877f2; color: white; padding: 14px 24px; text-decoration: none; border-radius: 8px; font-weight: 600; font-size: 16px; display: inline-block; transition: all 0.2s; box-shadow: 0 4px 6px -1px rgba(24, 119, 242, 0.3);">
                    <svg viewBox="0 0 24 24" width="24" height="24" style="vertical-align: middle; margin-right: 8px; fill: white;"><path d="M23.9981 11.9991C23.9981 5.37216 18.626 0 11.9991 0C5.37216 0 0 5.37216 0 11.9991C0 17.9882 4.38789 22.9522 10.1242 23.8524V15.4676H7.07758V11.9991H10.1242V9.35553C10.1242 6.34826 11.9156 4.68714 14.6564 4.68714C15.9692 4.68714 17.3424 4.92149 17.3424 4.92149V7.87439H15.8294C14.3388 7.87439 13.8739 8.79933 13.8739 9.74824V11.9991H17.2018L16.6698 15.4676H13.8739V23.8524C19.6103 22.9522 23.9981 17.9882 23.9981 11.9991Z"/></svg>
                    Đăng Nhập Bằng Facebook
                </a>
            </div>
        {% else %}
            <div style="text-align: right; margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px dashed var(--border-color);">
                <span style="color: var(--success); font-weight: 600; font-size: 14px;">✅ Đã cấp quyền Facebook thành công</span>
                <a href="/logout" style="margin-left: 10px; color: var(--error); text-decoration: none; font-size: 14px; font-weight: 500;">(Đăng xuất)</a>
            </div>
            <form method="POST" enctype="multipart/form-data">
                <div class="form-group">
                    <label>Tệp danh sách Excel chứa Link (.xlsx):</label>
                    <input type="file" name="excel_file" accept=".xlsx, .xls" required>
                </div>
                <div class="form-group">
                    <label>Hình ảnh đính kèm <span class="optional">(Không bắt buộc)</span>:</label>
                    <input type="file" name="image_file" accept="image/*">
                </div>
                <div class="form-group">
                    <label>Nội dung bài viết:</label>
                    <textarea name="message" required rows="5" placeholder="Nhập nội dung..."></textarea>
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
                            <strong>Mục tiêu: {{ res.group }}</strong> <br> 
                            <span style="font-weight: 400; font-size: 13px; margin-top: 4px; display: inline-block;">Trạng thái: {{ res.status }}</span>
                        </li>
                    {% endfor %}
                    </ul>
                </div>
            {% endif %}
        {% endif %}
    </div>
</body>
</html>
"""

def extract_target_id(input_str):
    input_str = str(input_str).strip()
    if input_str.isdigit(): return input_str
    id_match = re.search(r'id=(\d+)', input_str)
    if id_match: return id_match.group(1)
    group_match = re.search(r'groups/(\d+)', input_str)
    if group_match: return group_match.group(1)
    vanity_match = re.search(r'facebook\.com/([^/?]+)', input_str)
    if vanity_match:
        name = vanity_match.group(1)
        if name not in ['groups', 'profile.php', 'pages']: return name
    return None

# --- CÁC CỔNG GIAO TIẾP VỚI FACEBOOK ---
@app.route('/callback')
def callback():
    code = request.args.get('code')
    if not code:
        return "Lỗi: Bạn đã hủy quá trình đăng nhập hoặc Facebook từ chối kết nối.", 400
    
    # Mang mã code đi đổi lấy Access Token
    token_url = f"https://graph.facebook.com/v19.0/oauth/access_token?client_id={APP_ID}&redirect_uri={REDIRECT_URI}&client_secret={APP_SECRET}&code={code}"
    response = requests.get(token_url).json()
    
    if 'access_token' in response:
        # Lưu token vào két sắt (session)
        session['fb_access_token'] = response['access_token']
        return redirect(url_for('index'))
    else:
        return f"Lỗi không thể tạo mã Token: {response}", 400

@app.route('/logout')
def logout():
    session.pop('fb_access_token', None)
    return redirect(url_for('index'))

@app.route('/', methods=['GET', 'POST'])
def index():
    # Kiểm tra xem đã đăng nhập chưa
    if 'fb_access_token' not in session:
        scope = "publish_to_groups,pages_show_list,pages_read_engagement,pages_manage_posts"
        login_url = f"https://www.facebook.com/v19.0/dialog/oauth?client_id={APP_ID}&redirect_uri={REDIRECT_URI}&scope={scope}"
        return render_template_string(HTML_TEMPLATE, is_logged_in=False, login_url=login_url)

    # Nếu đã đăng nhập, tiến hành chạy form
    results = []
    if request.method == 'POST':
        master_token = session['fb_access_token']
        message = request.form['message']
        delay = int(request.form.get('delay', 5))
        excel_file = request.files.get('excel_file')
        image_file = request.files.get('image_file')

        image_bytes = None; image_name = None
        if image_file and image_file.filename != '':
            image_bytes = image_file.read()
            image_name = image_file.filename

        page_tokens = {}
        try:
            acc_url = f"https://graph.facebook.com/v19.0/me/accounts?access_token={master_token}"
            acc_res = requests.get(acc_url).json()
            if 'data' in acc_res:
                for page in acc_res['data']:
                    page_tokens[page['id']] = page['access_token']
        except:
            pass 

        if excel_file and excel_file.filename != '':
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
            excel_file.save(filepath)

            try:
                wb = openpyxl.load_workbook(filepath, data_only=True)
                sheet = wb.active
                target_list = []
                for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
                    if row[0]:
                        target_list.append(str(row[0]))
                wb.close()

                for item in target_list:
                    raw_id = extract_target_id(item)
                    if not raw_id:
                        results.append({'group': item, 'status': '❌ Không bóc tách được ID'})
                        continue

                    target_id = raw_id
                    if not str(raw_id).isdigit():
                        try:
                            resolve_url = f"https://graph.facebook.com/v19.0/{raw_id}?access_token={master_token}"
                            res_data = requests.get(resolve_url).json()
                            if 'id' in res_data:
                                target_id = res_data['id']
                        except:
                            pass
                    
                    active_token = page_tokens.get(target_id, master_token)

                    if image_bytes:
                        url = f"https://graph.facebook.com/v19.0/{target_id}/photos"
                        payload = {'message': message, 'access_token': active_token}
                        files = {'source': (image_name, image_bytes, 'image/jpeg')}
                        response = requests.post(url, data=payload, files=files)
                    else:
                        url = f"https://graph.facebook.com/v19.0/{target_id}/feed"
                        payload = {'message': message, 'access_token': active_token}
                        response = requests.post(url, data=payload)
                    
                    try:
                        data = response.json()
                        if 'id' in data or 'post_id' in data:
                            results.append({'group': raw_id, 'status': f"✅ Đăng thành công (Mã: {data.get('post_id', data.get('id'))})"})
                        else:
                            error_msg = data.get('error', {}).get('message', 'Lỗi từ chối')
                            results.append({'group': raw_id, 'status': f"❌ Bị từ chối: {error_msg}"})
                    except Exception as e:
                        results.append({'group': raw_id, 'status': f"❌ Lỗi kết nối: {str(e)}"})
                    
                    time.sleep(delay)

            except Exception as e:
                results.append({'group': 'Hệ thống', 'status': f"❌ Lỗi xử lý Excel: {str(e)}"})
            finally:
                if os.path.exists(filepath):
                    os.remove(filepath)
                
    return render_template_string(HTML_TEMPLATE, is_logged_in=True, results=results)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
