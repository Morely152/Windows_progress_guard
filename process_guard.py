import psutil
import subprocess
import os
import sys
import json
import time

def process_monitoring(process_route):
    """监测进程"""
    process_name = os.path.basename(process_route)
    for proc in psutil.process_iter(['name']):
        if process_name.lower() in proc.info['name'].lower():
            return True
        
def process_launch(process_route):
    """启动进程"""
    subprocess.Popen(process_route)

def file_path(file_name):
    if getattr(sys, 'frozen', False):  # 表示程序是打包后的可执行文件
        app_dir = os.path.dirname(sys.executable)
    else:  # 表示以脚本形式运行
        app_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(app_dir, file_name)


def process_guard():
    """进程守护"""
    while True:
        with open(file_path("tasks.json"),"r") as processes:
            tasks = json.load(processes)
        
        for process_route in list(tasks.keys()):
            if tasks[process_route] == True:
                if not process_monitoring(process_route):
                    process_launch(process_route)
        
        time.sleep(1)