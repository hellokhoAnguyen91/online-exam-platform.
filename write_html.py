import os

html_content = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Nền tảng thi trực tuyến</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 0; background: #f4f4f9; }
        .container { max-width: 800px; margin: 40px auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        h1, h2, h3 { color: #333; }
        input, button, select { display: block; margin: 10px 0; padding: 10px; width: 100%; box-sizing: border-box; }
        button { background: #007bff; color: white; border: none; cursor: pointer; }
        button:hover { background: #0056b3; }
        .hidden { display: none; }
        .question { margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 10px; }
        .options label { display: block; margin: 5px 0; }
        .top-bar { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #ccc; padding-bottom: 10px; margin-bottom: 20px; }
        #timer { font-size: 1.5em; font-weight: bold; color: red; }
    </style>
</head>
<body>

<div class="container" id="login-view">
    <h1>Đăng nhập</h1>
    <input type="text" id="username" placeholder="Tên đăng nhập (MSSV hoặc admin)">
    <input type="password" id="password" placeholder="Mật khẩu (Ngày sinh hoặc mật khẩu admin)">
    <button onclick="login()">Đăng nhập</button>
    <p id="login-error" style="color:red;"></p>
</div>

<div class="container hidden" id="admin-view">
    <div class="top-bar">
        <h2>Bảng điều khiển Quản trị viên</h2>
        <button onclick="logout()" style="width:auto;">Đăng xuất</button>
    </div>
    
    <h3>1. Cấu hình</h3>
    <label>Số lượng câu hỏi mỗi đề thi:</label>
    <input type="number" id="num_questions" value="10">
    <label>Thời gian thi (Phút):</label>
    <input type="number" id="duration_minutes" value="30">
    <label>Thời gian mở thi (vd: 2026-09-30T10:00:00) (Tùy chọn):</label>
    <input type="datetime-local" id="open_time">
    <label>Thời gian đóng thi (vd: 2026-09-30T12:00:00) (Tùy chọn):</label>
    <input type="datetime-local" id="close_time">
    <button onclick="saveConfig()">Lưu Cấu hình</button>
    
    <h3>2. Tải lên Ngân hàng câu hỏi (.docx)</h3>
    <input type="file" id="file_questions" accept=".docx">
    <button onclick="uploadQuestions()">Tải lên Câu hỏi</button>
    
    <h3>3. Tải lên Danh sách sinh viên (.xlsx)</h3>
    <p>Các cột: MSSV, FullName, DOB</p>
    <input type="file" id="file_students" accept=".xlsx">
    <button onclick="uploadStudents()">Tải lên Sinh viên</button>
    
    <h3>4. Xem Câu hỏi đã tải lên</h3>
    <button onclick="loadQuestions()">Xem Danh sách Câu hỏi</button>
    <div id="questions-list" style="margin-top:10px; max-height: 400px; overflow-y: auto; border: 1px solid #ccc; padding: 10px; display: none;"></div>

    <h3>5. Kết quả thi</h3>
    <button onclick="loadResults()">Tải Kết quả</button>
    <button onclick="exportResults()" style="background-color: #28a745;">Xuất Excel</button>
    <table id="results-table" border="1" width="100%" style="margin-top:10px; border-collapse: collapse;">
        <thead><tr><th>MSSV</th><th>Họ tên</th><th>Điểm số</th><th>Thời gian nộp</th></tr></thead>
        <tbody></tbody>
    </table>
</div>

<div class="container hidden" id="student-view">
    <div class="top-bar">
        <h2 id="student-name">Bài thi</h2>
        <div id="timer">--:--</div>
        <button onclick="logout()" style="width:auto;">Đăng xuất</button>
    </div>
    
    <div id="exam-content"></div>
    <button id="submit-exam-btn" onclick="submitExam()" class="hidden">Nộp Bài thi</button>
    
    <div id="result-view" class="hidden">
        <h2>Đã nộp bài thi</h2>
        <p>Điểm của bạn: <strong id="student-score"></strong></p>
    </div>
</div>

<script>
    let token = localStorage.getItem('token');
    let timerInterval = null;
    let answers = {};

    async function req(url, options = {}) {
        if(!options.headers) options.headers = {};
        if(token) options.headers['Authorization'] = 'Bearer ' + token;
        const res = await fetch(url, options);
        if(res.status === 401) { logout(); throw new Error('Không có quyền truy cập'); }
        return res;
    }

    async function login() {
        const u = document.getElementById('username').value;
        const p = document.getElementById('password').value;
        const formData = new URLSearchParams();
        formData.append('username', u);
        formData.append('password', p);
        
        const res = await fetch('/token', {
            method: 'POST',
            headers: {'Content-Type': 'application/x-www-form-urlencoded'},
            body: formData
        });
        if(res.ok) {
            const data = await res.json();
            token = data.access_token;
            localStorage.setItem('token', token);
            checkUser();
        } else {
            document.getElementById('login-error').innerText = "Tên đăng nhập hoặc mật khẩu không chính xác";
        }
    }

    function logout() {
        token = null;
        localStorage.removeItem('token');
        clearInterval(timerInterval);
        document.getElementById('login-view').classList.remove('hidden');
        document.getElementById('admin-view').classList.add('hidden');
        document.getElementById('student-view').classList.add('hidden');
    }

    async function checkUser() {
        if(!token) return;
        try {
            const res = await req('/api/me');
            if(res.ok) {
                const user = await res.json();
                document.getElementById('login-view').classList.add('hidden');
                if(user.is_admin) {
                    document.getElementById('admin-view').classList.remove('hidden');
                } else {
                    document.getElementById('student-view').classList.remove('hidden');
                    document.getElementById('student-name').innerText = "Bài thi: " + user.fullname;
                    loadExam();
                }
            }
        } catch(e) {}
    }

    async function saveConfig() {
        const formData = new FormData();
        formData.append('num_questions', document.getElementById('num_questions').value);
        formData.append('duration_minutes', document.getElementById('duration_minutes').value);
        if (document.getElementById('open_time').value) formData.append('open_time', document.getElementById('open_time').value);
        if (document.getElementById('close_time').value) formData.append('close_time', document.getElementById('close_time').value);
        const res = await req('/api/admin/config', { method: 'POST', body: formData });
        if(res.ok) alert('Đã lưu cấu hình');
    }

    async function uploadQuestions() {
        const file = document.getElementById('file_questions').files[0];
        if(!file) return alert('Vui lòng chọn tệp');
        const formData = new FormData();
        formData.append('file', file);
        const res = await req('/api/admin/upload_questions', { method: 'POST', body: formData });
        if(res.ok) alert('Đã tải lên câu hỏi');
    }

    async function uploadStudents() {
        const file = document.getElementById('file_students').files[0];
        if(!file) return alert('Vui lòng chọn tệp');
        const formData = new FormData();
        formData.append('file', file);
        const res = await req('/api/admin/upload_students', { method: 'POST', body: formData });
        if(res.ok) alert('Đã tải lên sinh viên');
    }

    async function loadQuestions() {
        const res = await req('/api/admin/questions');
        if(!res.ok) return alert('Lỗi khi tải câu hỏi');
        const data = await res.json();
        const container = document.getElementById('questions-list');
        container.style.display = 'block';
        container.innerHTML = `<h4>Tổng số câu hỏi: ${data.count}</h4>`;
        data.questions.forEach((q, idx) => {
            container.innerHTML += `<div style="border-bottom: 1px solid #eee; margin-bottom: 10px; padding-bottom: 10px;">
                <strong>Câu ${idx + 1}:</strong> ${q.content}<br>
                A: ${q.option_a}<br>
                B: ${q.option_b}<br>
                C: ${q.option_c}<br>
                D: ${q.option_d}<br>
                <em>Đáp án: ${q.correct_option}</em>
            </div>`;
        });
    }

    async function loadResults() {
        const res = await req('/api/admin/results');
        const data = await res.json();
        const tbody = document.querySelector('#results-table tbody');
        tbody.innerHTML = '';
        data.forEach(r => {
            tbody.innerHTML += `<tr><td>${r.username}</td><td>${r.fullname}</td><td>${r.score}</td><td>${r.submit_time || '-'}</td></tr>`;
        });
    }

    async function exportResults() {
        const res = await req('/api/admin/results/export');
        if(res.ok) {
            const blob = await res.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'results.xlsx';
            document.body.appendChild(a);
            a.click();
            a.remove();
        } else {
            alert('Lỗi xuất Excel');
        }
    }

    async function loadExam() {
        const res = await req('/api/exam');
        if(!res.ok) {
            const err = await res.json();
            alert(err.detail || 'Lỗi khi tải bài thi');
            return;
        }
        const data = await res.json();
        if(data.status === 'submitted') {
            showScore(data.score);
        } else {
            answers = data.answers || {};
            renderExam(data.questions);
            startTimer(data.time_left);
        }
    }

    function renderExam(questions) {
        const container = document.getElementById('exam-content');
        container.innerHTML = '';
        questions.forEach((q, idx) => {
            let html = `<div class="question"><p><strong>Câu ${idx+1}:</strong> ${q.content}</p><div class="options">`;
            ['a','b','c','d'].forEach(opt => {
                const key = 'option_' + opt;
                if(q[key]) {
                    const checked = answers[q.id] === opt.toUpperCase() ? 'checked' : '';
                    html += `<label><input type="radio" name="q_${q.id}" value="${opt.toUpperCase()}" onchange="saveAnswer(${q.id}, '${opt.toUpperCase()}')" ${checked}> ${q[key]}</label>`;
                }
            });
            html += `</div></div>`;
            container.innerHTML += html;
        });
        document.getElementById('submit-exam-btn').classList.remove('hidden');
    }

    async function saveAnswer(qId, val) {
        answers[qId] = val;
        await req('/api/exam/save_progress', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({answers})
        });
    }

    async function submitExam() {
        if(!confirm('Bạn có chắc chắn muốn nộp bài?')) return;
        forceSubmit();
    }
    
    async function forceSubmit() {
        clearInterval(timerInterval);
        const res = await req('/api/exam/submit', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({answers})
        });
        if(res.ok) {
            const data = await res.json();
            showScore(data.score);
        }
    }

    function showScore(score) {
        document.getElementById('exam-content').classList.add('hidden');
        document.getElementById('submit-exam-btn').classList.add('hidden');
        document.getElementById('timer').classList.add('hidden');
        document.getElementById('result-view').classList.remove('hidden');
        document.getElementById('student-score').innerText = score;
    }

    function startTimer(secondsLeft) {
        clearInterval(timerInterval);
        function update() {
            if(secondsLeft <= 0) {
                clearInterval(timerInterval);
                forceSubmit();
                return;
            }
            secondsLeft--;
            const m = Math.floor(secondsLeft / 60).toString().padStart(2, '0');
            const s = Math.floor(secondsLeft % 60).toString().padStart(2, '0');
            document.getElementById('timer').innerText = `${m}:${s}`;
        }
        update();
        timerInterval = setInterval(update, 1000);
    }

    if(token) checkUser();
</script>
</body>
</html>
"""

with open(r"C:\Users\WIN 10\.gemini\antigravity\scratch\online-exam-platform\static\index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("Done")
