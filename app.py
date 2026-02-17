#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公文格式调整工具 - Web版本
基于Flask实现的Web界面
"""

import os
import sys
import signal
import socket
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import tempfile
import shutil
from datetime import datetime

# 导入核心格式化函数
from gongwen_formatter_cli import format_document

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB 最大文件大小
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'docx'}

def check_and_kill_port(port):
    """检查端口是否被占用，如果占用则尝试释放"""
    try:
        # 检查端口是否被占用
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex(('localhost', port))
        sock.close()
        
        if result == 0:
            print(f"⚠️  端口 {port} 已被占用，尝试释放...")
            
            # macOS/Linux
            if sys.platform != 'win32':
                try:
                    import subprocess
                    # 查找占用端口的进程
                    result = subprocess.run(
                        ['lsof', '-ti', f':{port}'],
                        capture_output=True,
                        text=True
                    )
                    pids = result.stdout.strip().split('\n')
                    
                    # 终止这些进程
                    for pid in pids:
                        if pid:
                            try:
                                os.kill(int(pid), signal.SIGTERM)
                                print(f"  ✅ 已终止进程 {pid}")
                            except:
                                pass
                    
                    import time
                    time.sleep(1)
                    print(f"  ✅ 端口 {port} 已释放")
                    return True
                except Exception as e:
                    print(f"  ❌ 无法自动释放端口: {e}")
                    return False
            
            # Windows
            else:
                try:
                    import subprocess
                    # 查找占用端口的进程
                    result = subprocess.run(
                        ['netstat', '-ano', '-p', 'TCP'],
                        capture_output=True,
                        text=True
                    )
                    
                    for line in result.stdout.split('\n'):
                        if f':{port}' in line and 'LISTENING' in line:
                            parts = line.split()
                            pid = parts[-1]
                            try:
                                subprocess.run(['taskkill', '/F', '/PID', pid], check=True)
                                print(f"  ✅ 已终止进程 {pid}")
                            except:
                                pass
                    
                    import time
                    time.sleep(1)
                    print(f"  ✅ 端口 {port} 已释放")
                    return True
                except Exception as e:
                    print(f"  ❌ 无法自动释放端口: {e}")
                    return False
        else:
            return True
            
    except Exception as e:
        print(f"  ❌ 检查端口时出错: {e}")
        return True

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """首页"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传和格式化"""
    try:
        # 检查是否有文件
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '没有选择文件'}), 400
        
        file = request.files['file']
        
        # 检查文件名
        if file.filename == '':
            return jsonify({'success': False, 'error': '没有选择文件'}), 400
        
        # 检查文件类型
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': '只支持 .docx 格式的文件'}), 400
        
        # 保存上传的文件
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        temp_input = os.path.join(app.config['UPLOAD_FOLDER'], f'temp_{timestamp}_{filename}')
        file.save(temp_input)
        
        # 处理文档
        success = format_document(temp_input)
        
        if not success:
            os.remove(temp_input)
            return jsonify({'success': False, 'error': '文档处理失败，请检查文档格式'}), 500
        
        # 获取输出文件路径
        dir_name = os.path.dirname(temp_input)
        base_name = os.path.basename(temp_input)
        output_path = os.path.join(dir_name, f"done_{base_name}")
        
        # 检查输出文件是否存在
        if not os.path.exists(output_path):
            os.remove(temp_input)
            return jsonify({'success': False, 'error': '输出文件生成失败'}), 500
        
        # 读取输出文件
        with open(output_path, 'rb') as f:
            output_data = f.read()
        
        # 清理临时文件
        os.remove(temp_input)
        os.remove(output_path)
        
        # 保存处理后的文件到临时位置
        final_output = os.path.join(app.config['UPLOAD_FOLDER'], f'done_{timestamp}_{filename}')
        with open(final_output, 'wb') as f:
            f.write(output_data)
        
        # 返回文件下载链接
        return jsonify({
            'success': True,
            'download_url': f'/download/{os.path.basename(final_output)}',
            'filename': f'done_{filename}'
        })
        
    except Exception as e:
        print(f"处理错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'服务器错误: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """下载处理后的文件"""
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        if not os.path.exists(file_path):
            return jsonify({'success': False, 'error': '文件不存在'}), 404
        
        # 发送文件并在发送后删除
        response = send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
        # 设置一个回调来删除文件（Flask会在发送后执行）
        @response.call_on_close
        def cleanup():
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass
        
        return response
        
    except Exception as e:
        print(f"下载错误: {str(e)}")
        return jsonify({'success': False, 'error': f'下载失败: {str(e)}'}), 500

if __name__ == '__main__':
    PORT = 5000
    
    print("\n" + "=" * 60)
    print("  📄 公文格式调整工具 - Web版")
    print("=" * 60)
    
    # 检查并清理端口
    print("\n🔍 检查端口...")
    if check_and_kill_port(PORT):
        print("\n✅ 服务启动成功！")
        print(f"🌐 请在浏览器中访问: http://localhost:{PORT}")
        print("\n按 Ctrl+C 停止服务\n")
        print("=" * 60 + "\n")
        
        try:
            app.run(debug=True, host='0.0.0.0', port=PORT, use_reloader=False)
        except KeyboardInterrupt:
            print("\n\n👋 服务已停止\n")
    else:
        print(f"\n❌ 无法启动服务，端口 {PORT} 被占用")
        print(f"请手动关闭占用端口的程序，或修改 app.py 中的端口号\n")
