import os
import sys
import json
import ctypes
import threading
import winreg as reg
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import *
from process_guard import process_guard

class TaskWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.auto_minimize_to_tray = json.load(open(self.file_path("settings.json")))['auto_minimize_to_tray']
        self.setWindowIcon(QIcon(self.file_path("icon.ico")))
        self.on_app_started()
        self.tray_icon.activated.connect(self.on_tray_activated)
        
    def initUI(self):
        """初始化主窗口"""

        # 创建按钮
        self.addTaskButton = QPushButton('添加任务', self)
        self.delTaskButton = QPushButton('删除任务', self)
        self.startGuardButton = QPushButton('开启守护', self)
        self.stopGuardButton = QPushButton('停止守护', self)
        self.settingsButton = QPushButton('软件设置', self)

        # 按钮大小
        self.addTaskButton.setFixedSize(100, 35)
        self.delTaskButton.setFixedSize(100, 35)
        self.startGuardButton.setFixedSize(100, 35)
        self.stopGuardButton.setFixedSize(100, 35)
        self.settingsButton.setFixedSize(100, 35)

        
        # 设置按钮事件连接（根据需要添加功能）
        self.addTaskButton.clicked.connect(self.add_task)
        self.delTaskButton.clicked.connect(self.del_task)
        self.startGuardButton.clicked.connect(self.start_guard)
        self.stopGuardButton.clicked.connect(self.stop_guard)
        self.settingsButton.clicked.connect(self.settings)
        
        # 创建表格
        self.table = QTableWidget(10, 3, self)  # 10行3列
        # 设置表头
        self.table.setHorizontalHeaderLabels(['名称', '状态', '路径'])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)  # 不可编辑
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)  # 只能选中整行
        
         # 设置表格列宽
        self.table.setColumnWidth(0, 150)
        self.table.setColumnWidth(1, 150)
        self.table.setColumnWidth(2, 280)
        
        # 创建布局
        buttonLayout = QHBoxLayout()
        buttonLayout.addWidget(self.addTaskButton)
        buttonLayout.addWidget(self.delTaskButton)
        buttonLayout.addWidget(self.startGuardButton)
        buttonLayout.addWidget(self.stopGuardButton)
        buttonLayout.addWidget(self.settingsButton)
        
        mainLayout = QVBoxLayout()
        mainLayout.addLayout(buttonLayout)
        mainLayout.addWidget(self.table)
        
        self.setLayout(mainLayout)
        self.setWindowTitle('进程守护')
        self.setFixedSize(650,700)
        self.show()
        self.load_tasks_from_json()

        # 创建托盘图标
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(self.file_path("icon.ico")))

    # 推算目标文件的路径
    def file_path(self, file_name):
            if getattr(sys, 'frozen', False):  # 表示程序是打包后的可执行文件
                app_dir = os.path.dirname(sys.executable)
            else:  # 表示以脚本形式运行
                app_dir = os.path.dirname(os.path.abspath(__file__))
            return os.path.join(app_dir, file_name)


    def load_tasks_from_json(self):
        # 尝试读取tasks.json文件，并将任务名称添加到列表中
        try:
            with open(self.file_path("tasks.json"), 'r') as f:
                tasks_dict = json.load(f)
                row = 0  # 从第一行开始填充数据
                for task_path, status in tasks_dict.items():
                    # 提取.exe文件名作为任务名称
                    task_name = os.path.basename(task_path)
                    # 在“名称”一列中添加任务名称
                    name_item = QTableWidgetItem(task_name)
                    self.table.setItem(row, 0, name_item)  # 设置名称列
                    
                    # 根据状态设置“状态”列的内容
                    if status:
                        status_item = QTableWidgetItem("已启用")
                    else:
                        status_item = QTableWidgetItem("未启用")
                    self.table.setItem(row, 1, status_item)  # 设置状态列
                    
                    # 在“路径”一列中添加任务路径
                    path_item = QTableWidgetItem(task_path)
                    self.table.setItem(row, 2, path_item)  # 设置路径列
                    
                    row += 1  # 移到下一行
                    if row >= self.table.rowCount():
                        break  # 如果行数超过表格行数则停止添加
        except FileNotFoundError:
            print("tasks.json文件未找到。")
        except json.JSONDecodeError:
            print("tasks.json文件格式错误。")

    def add_task(self):
        # 创建子窗口
        dialog = QDialog(self)
        dialog.setWindowTitle('添加任务')

        # 设置子窗口大小
        dialog.setFixedSize(450, 150)
        
        # 创建布局和控件
        mainLayout = QVBoxLayout(dialog)  # 创建主垂直布局
        
        # 创建水平布局用于放置文本框和按钮
        hLayout = QHBoxLayout()  # 创建水平布局
        self.filePathLineEdit = QLineEdit(dialog)  # 创建文本框
        self.filePathLineEdit.setReadOnly(True)  # 设置为只读，防止用户手动编辑
        hLayout.addWidget(self.filePathLineEdit)  # 将文本框添加到水平布局中
        
        self.chooseFileButton = QPushButton('选择文件', dialog)  # 注意按钮文本改为“选择文件”
        self.chooseFileButton.clicked.connect(self.choose_exe_file)  # 连接点击事件到choose_exe_file方法
        hLayout.addWidget(self.chooseFileButton)  # 将按钮添加到水平布局中
        
        # 添加水平布局到主垂直布局中
        mainLayout.addLayout(hLayout)
        
        # 添加确认按钮
        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok, dialog)
        buttonBox.accepted.connect(dialog.accept)  # 点击Ok时关闭对话框
        mainLayout.addWidget(buttonBox)  # 将确认按钮添加到主垂直布局中
        
        # 显示子窗口
        if dialog.exec_() == QDialog.Accepted:  # 使用exec_会阻塞主窗口，直到对话框关闭
            file_path = self.filePathLineEdit.text()  # 获取文件路径
            self.save_file_path(file_path)  # 调用新方法来保存文件路径
        self.refresh_table()  # 刷新表格
        
    def refresh_table(self):
        # 清除表格中的所有内容，但保留表头和行数
        self.table.clearContents()
    
        # 重新加载任务数据并填充到表格中
        self.load_tasks_from_json()

    
    def del_task(self):
        # 获取选中的行（如果有的话）
        selected_rows = self.table.selectionModel().selectedRows()
        if selected_rows:
            # 假设我们只处理单行选择，取第一行
            selected_row = selected_rows[0].row()
            
            # 获取该行的路径列数据
            path_item = self.table.item(selected_row, 2)  # 假设路径在第三列
            if path_item:
                path_to_delete = path_item.text()
                
                # 从tasks.json中删除对应的键值对
                try:
                    with open(self.file_path("tasks.json"), 'r') as f:
                        tasks_dict = json.load(f)
                    
                    if path_to_delete in tasks_dict:
                        del tasks_dict[path_to_delete]
                        
                        # 将更新后的字典写回json文件
                        with open(self.file_path("tasks.json"), 'w') as f:
                            json.dump(tasks_dict, f, ensure_ascii=False, indent=4)
                        
                        # 刷新表格
                        self.refresh_table()
                        
                        print(f"已删除任务: {path_to_delete}")
                    else:
                        print("选中的任务不在任务列表中。")
                except FileNotFoundError:
                    print("tasks.json文件未找到。")
                except json.JSONDecodeError:
                    print("tasks.json文件格式错误。")
            else:
                print("未选中有效的任务。")
        else:
            print("请至少选中一个任务以进行删除。")        

    def start_guard(self):
        # 获取选中的行（如果有的话）
        selected_rows = self.table.selectionModel().selectedRows()
        if selected_rows:
            # 假设我们只处理单行选择，取第一行
            selected_row = selected_rows[0].row()
            
            # 获取该行的路径列数据
            path_item = self.table.item(selected_row, 2)  # 假设路径在第三列
            if path_item:
                path_to_enable = path_item.text()
                
                # 修改tasks.json中对应的键的值为True
                try:
                    with open(self.file_path("tasks.json"), 'r') as f:
                        tasks_dict = json.load(f)
                    
                    if path_to_enable in tasks_dict:
                        tasks_dict[path_to_enable] = True
                        
                        # 将更新后的字典写回json文件
                        with open(self.file_path("tasks.json"), 'w') as f:
                            json.dump(tasks_dict, f, ensure_ascii=False, indent=4)
                        
                        # 刷新表格
                        self.refresh_table()
                        
                        print(f"已启用守护: {path_to_enable}")
                    else:
                        print("选中的任务不在任务列表中。")
                except FileNotFoundError:
                    print("tasks.json文件未找到。")
                except json.JSONDecodeError:
                    print("tasks.json文件格式错误。")
            else:
                print("未选中有效的任务。")
        else:
            print("请至少选中一个任务以进行守护。")
        self.refresh_table()
        
    def stop_guard(self):
       # 获取选中的行（如果有的话）
        selected_rows = self.table.selectionModel().selectedRows()
        if selected_rows:
            # 假设我们只处理单行选择，取第一行
            selected_row = selected_rows[0].row()
            
            # 获取该行的路径列数据
            path_item = self.table.item(selected_row, 2)  # 假设路径在第三列
            if path_item:
                path_to_enable = path_item.text()
                
                # 修改tasks.json中对应的键的值为True
                try:
                    with open(self.file_path("tasks.json"), 'r') as f:
                        tasks_dict = json.load(f)
                    
                    if path_to_enable in tasks_dict:
                        tasks_dict[path_to_enable] = False
                        
                        # 将更新后的字典写回json文件
                        with open(self.file_path("tasks.json"), 'w') as f:
                            json.dump(tasks_dict, f, ensure_ascii=False, indent=4)
                        
                        # 刷新表格
                        self.refresh_table()
                        
                        print(f"已停止守护: {path_to_enable}")
                    else:
                        print("选中的任务不在任务列表中。")
                except FileNotFoundError:
                    print("tasks.json文件未找到。")
                except json.JSONDecodeError:
                    print("tasks.json文件格式错误。")
            else:
                print("未选中有效的任务。")
        else:
            print("请至少选中一个任务以停止守护。")
        self.refresh_table()

    def settings(self):
        # 创建一个子窗口
        self.settings_window = QDialog(self)
        self.settings_window.setWindowTitle('软件设置')
        self.settings_window.resize(300, 200)

        # 使用垂直布局来组织子窗口的UI元素
        layout = QVBoxLayout()

        # 创建一个复选框用于控制开机自启动
        self.autostart_checkbox = QCheckBox("开机自启动")
        layout.addWidget(self.autostart_checkbox)

        # 创建一个复选框用于控制缩小到托盘
        self.auto_minimize_to_tray_checkbox = QCheckBox("启动时最小化到托盘")
        self.auto_minimize_to_tray_checkbox.stateChanged.connect(self.on_minimize_to_tray_checkebox_changed)
        layout.addWidget(self.auto_minimize_to_tray_checkbox)

        # 创建一个按钮用于缩小到托盘
        minimize_to_tray_button = QPushButton("最小化到托盘")
        minimize_to_tray_button.clicked.connect(lambda: self.minimize_to_tray())
        layout.addWidget(minimize_to_tray_button)

        # 创建一个按钮用于保存设置
        save_button = QPushButton("保存设置")
        save_button.clicked.connect(self.save_settings)
        layout.addWidget(save_button)

        # 读取settings.json文件并设置复选框的状态
        try:
            with open(self.file_path("settings.json"), 'r') as f:
                settings_dict = json.load(f)
            if 'auto_start' in settings_dict:
                self.autostart_checkbox.setChecked(settings_dict['auto_start'])
            if 'auto_minimize_to_tray' in settings_dict:
                self.auto_minimize_to_tray_checkbox.setChecked(settings_dict['auto_minimize_to_tray'])
        except FileNotFoundError:
            print("settings.json文件未找到。")
        except json.JSONDecodeError:
           print("settings.json文件格式错误。")
        
        # 设置布局并显示子窗口
        self.settings_window.setLayout(layout)
        self.settings_window.exec_()
    
    def choose_exe_file(self):
        # 打开文件选择对话框，并限制只选择exe文件
        file_name, _ = QFileDialog.getOpenFileName(self, '选择EXE文件', '.', 'Executables (*.exe)')
        if file_name:
            print(f'选择的EXE文件是: {file_name}')
            self.filePathLineEdit.setText(file_name)  # 将选择的文件路径设置到文本框中显示
    
    def save_file_path(self, file_path):
        # 保存文件路径到json文件中
        save_file = self.file_path("tasks.json")  # 指定保存文件路径的文件名
        
        # 读取现有的json数据（如果有的话）
        try:
            with open(save_file, 'r') as f:
                tasks_dict = json.load(f)
        except FileNotFoundError:
            tasks_dict = {}
        
        # 添加新路径到字典中，值为True
        tasks_dict[file_path] = True
        
        # 将更新后的字典写回json文件
        with open(save_file, 'w') as f:
            json.dump(tasks_dict, f, ensure_ascii=False, indent=4)
    
    def save_settings(self):
        if self.autostart_checkbox.isChecked():
            try:
                # 打开注册表项
                key = reg.OpenKey(reg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, reg.KEY_SET_VALUE)
                # 设置注册表值
                reg.SetValueEx(key, "progress_guard", 0, reg.REG_SZ, os.path.abspath(sys.executable))
                reg.CloseKey(key)
                
                # 修改settings.json中对应的键的值为True
                with open(self.file_path("settings.json"), 'r') as f:
                    settings_dict = json.load(f)
                
                if 'auto_start' in settings_dict:
                    settings_dict['auto_start'] = True
                    
                    # 将更新后的字典写回json文件
                    with open(self.file_path("settings.json"), 'w') as f:
                        json.dump(settings_dict, f, ensure_ascii=False, indent=4)
            except Exception as e:
                print(f"Failed to set autostart: {e}")
        else:
            try:
                # 打开注册表项
                key = reg.OpenKey(reg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, reg.KEY_SET_VALUE)
                # 删除注册表值
                reg.DeleteValue(key, "progress_guard")
                reg.CloseKey(key)
                
                # 修改settings.json中对应的键的值为False
                with open(self.file_path("settings.json"), 'r') as f:
                    settings_dict = json.load(f)
                
                if 'auto_start' in settings_dict:
                    settings_dict['auto_start'] = False
                    
                    # 将更新后的字典写回json文件
                    with open(self.file_path("settings.json"), 'w') as f:
                        json.dump(settings_dict, f, ensure_ascii=False, indent=4)
            except FileNotFoundError:
                # 键值不存在时，忽略
                pass
            except Exception as e:
                print(f"Failed to remove autostart: {e}")
        
        # 关闭子窗口
        self.settings_window.close()

    def minimize_to_tray(self):
        for widget in QApplication.topLevelWidgets():
            if isinstance(widget,QDialog):
                widget.hide()
        self.hide()
        self.tray_icon.show()

    def show_normal(self):
        self.show()
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        self.activateWindow()
    
    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.show_normal()
    
    def on_minimize_to_tray_checkebox_changed(self,state):
        if state == Qt.Checked:
            self.auto_minimize_to_tray = True
            # 修改settings.json中对应的键的值为false
            with open(self.file_path("settings.json"), 'r') as f:
                settings_dict = json.load(f)
            if 'auto_minimize_to_tray' in settings_dict:
                settings_dict['auto_minimize_to_tray'] = True
                # 将更新后的字典写回json文件
                with open(self.file_path("settings.json"), 'w') as f:
                    json.dump(settings_dict, f, ensure_ascii=False, indent=4)
        else:
            self.auto_minimize_to_tray = False
            # 修改settings.json中对应的键的值为false
            with open(self.file_path("settings.json"), 'r') as f:
                settings_dict = json.load(f)
            if 'auto_minimize_to_tray' in settings_dict:
                settings_dict['auto_minimize_to_tray'] = False
                # 将更新后的字典写回json文件
                with open(self.file_path("settings.json"), 'w') as f:
                    json.dump(settings_dict, f, ensure_ascii=False, indent=4)
    
    def on_app_started(self):
        if self.auto_minimize_to_tray:
            self.minimize_to_tray()
        else:
            self.show_normal()

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def restart_as_admin():
    # Python 可执行文件
    executable = sys.executable
    # 当前脚本文件
    script = os.path.abspath(__file__)

    # 调用 ShellExecuteEx 提升权限
    params = f'cmd /c {executable} {script}'
    ret = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", executable, script, None, 1)
    return ret > 32

if __name__ == '__main__':
    # 检查是否以管理员权限运行
    if not is_admin():
        # 提升权限
        if restart_as_admin():
            sys.exit()
        else:
            print("Failed to elevate to administrator level")
    else:
        # 已经是管理员权限，继续执行后续代码
        print("Administrator privileges granted")

        app = QApplication(sys.argv)
        ex = TaskWindow()
        thread = threading.Thread(target=process_guard)
        app.aboutToQuit.connect(ex.on_app_started)
        thread.daemon = True
        thread.start()
        sys.exit(app.exec_())