"""
ChatBI 启动器 - 用于打包成exe
"""
import os
import sys
import webbrowser
import threading
import time

# 获取程序运行路径（支持打包后的路径）
if getattr(sys, 'frozen', False):
    # 打包后的路径
    base_path = sys._MEIPASS
else:
    # 开发环境路径
    base_path = os.path.dirname(os.path.abspath(__file__))

os.chdir(base_path)

def open_browser():
    """延迟打开浏览器"""
    time.sleep(3)
    webbrowser.open('http://127.0.0.1:5000')

if __name__ == '__main__':
    print("=" * 60)
    print("ChatBI 智能数据分析助手")
    print("=" * 60)
    print("正在启动服务...")
    print("")

    # 在新线程中打开浏览器
    threading.Thread(target=open_browser, daemon=True).start()

    # 导入并运行Flask应用
    from app import app, load_data

    # 加载数据
    load_data()

    print("")
    print("=" * 60)
    print("服务已启动！浏览器将自动打开")
    print("如果浏览器没有自动打开，请手动访问: http://127.0.0.1:5000")
    print("=" * 60)
    print("按 Ctrl+C 可停止服务")
    print("")

    # 运行Flask（不使用debug模式，避免多进程问题）
    app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
