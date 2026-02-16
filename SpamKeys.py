import sys
import json
import os
import time
import random
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel,  QSpinBox, QComboBox, QPushButton, QTextEdit,
    QGroupBox, QMessageBox, QFormLayout, QProgressBar
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

class SendThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    
    def __init__(self, message, iterations, speed, send_method):
        super().__init__()
        self.message = message
        self.iterations = iterations
        self.speed = speed
        self.send_method = send_method
    
    def run(self):
        import ctypes
        import time
        
        # 计算每条消息之间的延迟
        sleep_time = 1.0 / self.speed
        
        for i in range(1, self.iterations + 1):
            # 生成随机字符串
            random_str = ""
            for j in range(5):
                if random.random() > 0.5:
                    random_str += chr(97 + int(random.random() * 25))
                else:
                    random_str += str(int(random.random() * 9))
            
            # 替换消息模板中的随机字符串
            current_message = self.message.replace("{random}", random_str)
            
            # 发送消息内容
            self.send_keys(current_message)
            
            # 发送回车键或Ctrl+Enter
            if self.send_method == "1":
                self.send_enter()
            else:
                self.send_ctrl_enter()
            
            # 短暂延迟，确保系统有时间处理输入
            time.sleep(sleep_time)
            
            # 每10条消息更新一次进度，减少GUI更新频率
            if i % 10 == 0 or i == self.iterations:
                self.progress.emit(int((i / self.iterations) * 100))
                # 短暂暂停，让GUI有时间处理事件
                time.sleep(0.01)
        
        # 发送完成信号
        self.finished.emit(f"消息发送完成！共发送 {self.iterations} 条消息。")
    
    def send_keys(self, text):
        # 使用keybd_event函数发送文本
        import ctypes
        user32 = ctypes.windll.user32
        
        # 对于每个字符，使用keybd_event发送
        for char in text:
            # 对于ASCII字符，使用虚拟键码
            if 0x20 <= ord(char) <= 0x7E:
                vk = self.get_virtual_key(char)
                if vk:
                    # 检查是否需要按住Shift键
                    shift = False
                    if char.isupper() or char in "!@#$%^&*()_+{}|:\"<>?~`":
                        shift = True
                    
                    if shift:
                        user32.keybd_event(0x10, 0, 0, 0)  # VK_SHIFT
                    
                    user32.keybd_event(vk, 0, 0, 0)  # 按下按键
                    user32.keybd_event(vk, 0, 2, 0)  # 释放按键
                    
                    if shift:
                        user32.keybd_event(0x10, 0, 2, 0)  # 释放Shift
            else:
                # 对于非ASCII字符，使用SendInput函数
                try:
                    import ctypes
                    from ctypes import wintypes
                    
                    # 定义INPUT结构
                    class INPUT(ctypes.Structure):
                        _fields_ = [
                            ('type', wintypes.DWORD),
                            ('ki', ctypes.c_void_p)
                        ]
                    
                    class KEYBDINPUT(ctypes.Structure):
                        _fields_ = [
                            ('wVk', wintypes.WORD),
                            ('wScan', wintypes.WORD),
                            ('dwFlags', wintypes.DWORD),
                            ('time', wintypes.DWORD),
                            ('dwExtraInfo', ctypes.c_void_p)
                        ]
                    
                    # 常量定义
                    INPUT_KEYBOARD = 1
                    KEYEVENTF_KEYUP = 0x0002
                    KEYEVENTF_UNICODE = 0x0004
                    
                    # 创建KEYBDINPUT结构
                    ki = KEYBDINPUT(0, ord(char), KEYEVENTF_UNICODE, 0, None)
                    
                    # 创建INPUT结构数组
                    inputs = (INPUT * 2)()
                    
                    # 按下事件
                    inputs[0].type = INPUT_KEYBOARD
                    inputs[0].ki = ctypes.cast(ctypes.pointer(ki), ctypes.c_void_p)
                    
                    # 释放事件
                    ki.dwFlags |= KEYEVENTF_KEYUP
                    inputs[1].type = INPUT_KEYBOARD
                    inputs[1].ki = ctypes.cast(ctypes.pointer(ki), ctypes.c_void_p)
                    
                    # 发送输入
                    user32 = ctypes.windll.user32
                    user32.SendInput(2, ctypes.byref(inputs), ctypes.sizeof(INPUT))
                except Exception:
                    # 如果SendInput失败，跳过该字符
                    pass
    
    def send_enter(self):
        # 发送回车键
        import ctypes
        user32 = ctypes.windll.user32
        
        # 常量定义
        VK_RETURN = 0x0D
        
        # 按下回车键
        user32.keybd_event(VK_RETURN, 0, 0, 0)
        # 释放回车键
        user32.keybd_event(VK_RETURN, 0, 2, 0)
    
    def send_ctrl_enter(self):
        # 发送Ctrl+Enter
        import ctypes
        user32 = ctypes.windll.user32
        
        # 常量定义
        VK_CONTROL = 0x11
        VK_RETURN = 0x0D
        
        # 按下Ctrl
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        # 按下Enter
        user32.keybd_event(VK_RETURN, 0, 0, 0)
        # 释放Enter
        user32.keybd_event(VK_RETURN, 0, 2, 0)
        # 释放Ctrl
        user32.keybd_event(VK_CONTROL, 0, 2, 0)
    
    def get_virtual_key(self, char):
        # 简单的虚拟键映射
        key_map = {
            'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45,
            'f': 0x46, 'g': 0x47, 'h': 0x48, 'i': 0x49, 'j': 0x4A,
            'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E, 'o': 0x4F,
            'p': 0x50, 'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54,
            'u': 0x55, 'v': 0x56, 'w': 0x57, 'x': 0x58, 'y': 0x59,
            'z': 0x5A, '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33,
            '4': 0x34, '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38,
            '9': 0x39, ' ': 0x20, '!': 0x31, '@': 0x32, '#': 0x33,
            '$': 0x34, '%': 0x35, '^': 0x36, '&': 0x37, '*': 0x38,
            '(': 0x39, ')': 0x30, '-': 0xBD, '_': 0xBD, '=': 0xBB,
            '+': 0xBB, '[': 0xDB, '{': 0xDB, ']': 0xDD, '}': 0xDD,
            '\\': 0xDC, '|': 0xDC, ';': 0xBA, ':': 0xBA, "'": 0xDE,
            '"': 0xDE, ',': 0xBC, '<': 0xBC, '.': 0xBE, '>': 0xBE,
            '/': 0xBF, '?': 0xBF
        }
        return key_map.get(char.lower())

class SpamKeysApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SpamKeys - 消息发送工具")
        self.setGeometry(100, 100, 600, 500)
        self.setMinimumSize(500, 400)
        
        self.config_dir = os.path.join(os.environ['APPDATA'], 'Qxyz17', 'SpamKeys')
        self.config_file = os.path.join(self.config_dir, 'config.json')
        
        self.init_ui()
        self.load_config()
    
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题
        title_label = QLabel("SpamKeys - 消息发送工具")
        title_label.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 配置组
        config_group = QGroupBox("配置选项")
        config_layout = QFormLayout()
        config_group.setLayout(config_layout)
        
        # 消息内容
        self.message_edit = QTextEdit()
        self.message_edit.setPlaceholderText("请输入要发送的消息内容（使用{random}表示随机字符串）")
        self.message_edit.setFixedHeight(100)
        config_layout.addRow(QLabel("消息内容:"), self.message_edit)
        
        # 发送数量
        self.iterations_spin = QSpinBox()
        self.iterations_spin.setRange(1, 10000)
        self.iterations_spin.setValue(10)
        config_layout.addRow(QLabel("发送数量:"), self.iterations_spin)
        
        # 发送速度
        self.speed_spin = QSpinBox()
        self.speed_spin.setRange(1, 1000)
        self.speed_spin.setValue(10)
        config_layout.addRow(QLabel("发送速度 (条/秒):"), self.speed_spin)
        
        # 发送方式
        self.send_method_combo = QComboBox()
        self.send_method_combo.addItems(["按Enter发送", "按Ctrl+Enter发送"])
        config_layout.addRow(QLabel("发送方式:"), self.send_method_combo)
        
        main_layout.addWidget(config_group)
        
        # 操作按钮
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("开始发送")
        self.start_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.start_button.clicked.connect(self.start_sending)
        
        button_layout.addWidget(self.start_button)
        
        main_layout.addLayout(button_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # 状态信息
        self.status_label = QLabel("就绪")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setFont(QFont("Arial", 10))
        main_layout.addWidget(self.status_label)
    
    def start_sending(self):
        message = self.message_edit.toPlainText()
        iterations = self.iterations_spin.value()
        speed = self.speed_spin.value()
        send_method = str(self.send_method_combo.currentIndex() + 1)
        
        if not message:
            QMessageBox.warning(self, "警告", "请输入消息内容！")
            return
        
        # 显示确认对话框
        confirm_message = f"设置摘要：\n"
        confirm_message += f"消息模板: {message}\n"
        confirm_message += f"发送条数: {iterations}\n"
        confirm_message += f"发送速度: {speed}条/秒\n"
        confirm_message += f"发送方式: {self.send_method_combo.currentText()}\n\n"
        confirm_message += f"是否开始执行？"
        
        reply = QMessageBox.question(
            self, "确认设置", confirm_message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.No:
            return
        
        # 显示准备提示
        QMessageBox.information(
            self, "提示", "脚本将在5秒后开始，请确保目标窗口已激活且光标在输入框中！\n按确定后请快速切换到目标窗口。"
        )
        
        # 开始发送
        self.start_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.status_label.setText("准备中...")
        
        # 在后台线程中执行倒计时和发送
        import threading
        def countdown_and_send():
            # 倒计时
            for i in range(5, 0, -1):
                self.status_label.setText(f"准备中... {i}")
                QApplication.processEvents()
                time.sleep(1)
            
            self.status_label.setText("发送中...")
            
            # 开始发送
            self.send_thread = SendThread(message, iterations, speed, send_method)
            self.send_thread.progress.connect(self.update_progress)
            self.send_thread.finished.connect(self.send_finished)
            self.send_thread.start()
        
        # 创建并启动后台线程
        countdown_thread = threading.Thread(target=countdown_and_send)
        countdown_thread.daemon = True
        countdown_thread.start()
    
    def update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def send_finished(self, message):
        self.progress_bar.setVisible(False)
        self.start_button.setEnabled(True)
        self.status_label.setText("就绪")
        QMessageBox.information(self, "完成", message)
    
    def save_config(self):
        config = {
            "message": self.message_edit.toPlainText(),
            "iterations": self.iterations_spin.value(),
            "speed": self.speed_spin.value(),
            "send_method": self.send_method_combo.currentIndex()
        }
        
        # 确保配置目录存在
        os.makedirs(self.config_dir, exist_ok=True)
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            pass
    
    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                self.message_edit.setPlainText(config.get("message", ""))
                self.iterations_spin.setValue(config.get("iterations", 10))
                self.speed_spin.setValue(config.get("speed", 10))
                self.send_method_combo.setCurrentIndex(config.get("send_method", 0))
        except Exception as e:
            pass
    
    def closeEvent(self, event):
        # 程序退出时自动保存配置
        self.save_config()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    window = SpamKeysApp()
    window.show()
    
    sys.exit(app.exec())