# -*- coding: utf-8 -*-
"""
增强版邮件提醒程序 - 自动开始监控
功能：收到邮件后循环播放音乐，手动确认后停止，支持配置保存和自动开始监控
"""

import os
import sys
import time
import threading
import winsound
import json
import imaplib
import email
import hashlib
from email.header import decode_header
from tkinter import Tk, Label, Button, Frame, messagebox, filedialog, Entry, StringVar, Scrollbar, Text, Checkbutton, BooleanVar
from datetime import datetime
import psutil

class EnhancedEmailAlert:
    def __init__(self):
        self.is_running = False
        self.is_alerting = False
        self.sound_file = "alert.wav"
        self.processed_emails = set()
        self.processed_file = "processed_emails.json"
        self.config_file = "app_config.json"
        
        # 默认配置
        self.default_config = {
            "email_settings": {
                "server": "imap.qq.com",
                "port": "993",
                "email": "",
                "password": ""
            },
            "alert_settings": {
                "sound_file": "alert.wav",
                "alert_mode": "popup",
                "check_interval": 30,
                "auto_start": False  # 新增：是否自动开始监控
            },
            "window_position": {
                "width": 600,
                "height": 500
            }
        }
        
        # 加载配置和已处理的邮件记录
        self.load_config()
        self.load_processed_emails()
        
        self.create_ui()
        
        # 如果配置了自动开始，则启动监控
        if self.config["alert_settings"]["auto_start"]:
            self.auto_start_monitoring()
    
    def load_config(self):
        """加载应用程序配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                self.config = self.default_config.copy()
                self.save_config()
        except Exception as e:
            print(f"加载配置失败: {e}")
            self.config = self.default_config.copy()
    
    def save_config(self):
        """保存应用程序配置"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")
    
    def load_processed_emails(self):
        """加载已处理的邮件记录"""
        try:
            if os.path.exists(self.processed_file):
                with open(self.processed_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.processed_emails = set(data.get('processed_emails', []))
                    
                    if len(self.processed_emails) > 1000:
                        self.processed_emails = set(list(self.processed_emails)[-1000:])
        except Exception as e:
            print(f"加载已处理邮件记录失败: {e}")
            self.processed_emails = set()
    
    def save_processed_emails(self):
        """保存已处理的邮件记录"""
        try:
            with open(self.processed_file, 'w', encoding='utf-8') as f:
                json.dump({'processed_emails': list(self.processed_emails)}, f, ensure_ascii=False)
        except Exception as e:
            print(f"保存已处理邮件记录失败: {e}")
    
    def get_email_hash(self, email_data):
        """生成邮件的唯一标识哈希"""
        content = f"{email_data.get('from', '')}_{email_data.get('subject', '')}_{email_data.get('date', '')}"
        return hashlib.md5(content.encode('utf-8')).hexdigest()
    
    def create_ui(self):
        """创建用户界面"""
        self.root = Tk()
        self.root.title("增强版邮件提醒程序 - 自动开始监控")
        self.root.geometry("600x600")
        self.root.resizable(True, True)
        
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 邮箱设置框架
        email_frame = Frame(self.root, padx=10, pady=10)
        email_frame.pack(fill="x")
        
        Label(email_frame, text="邮箱设置", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w", columnspan=4, pady=(0,10))
        
        Label(email_frame, text="IMAP服务器:").grid(row=1, column=0, sticky="w")
        self.server_var = StringVar(value=self.config["email_settings"]["server"])
        Entry(email_frame, textvariable=self.server_var, width=20).grid(row=1, column=1, padx=5)
        
        Label(email_frame, text="端口:").grid(row=1, column=2, sticky="w")
        self.port_var = StringVar(value=self.config["email_settings"]["port"])
        Entry(email_frame, textvariable=self.port_var, width=8).grid(row=1, column=3, padx=5)
        
        Label(email_frame, text="邮箱地址:").grid(row=2, column=0, sticky="w", pady=5)
        self.email_var = StringVar(value=self.config["email_settings"]["email"])
        Entry(email_frame, textvariable=self.email_var, width=30).grid(row=2, column=1, columnspan=3, padx=5, sticky="w")
        
        Label(email_frame, text="密码/授权码:").grid(row=3, column=0, sticky="w", pady=5)
        self.password_var = StringVar(value=self.config["email_settings"]["password"])
        Entry(email_frame, textvariable=self.password_var, show="*", width=30).grid(row=3, column=1, columnspan=3, padx=5, sticky="w")
        
        # 声音设置框架
        sound_frame = Frame(self.root, padx=10, pady=10)
        sound_frame.pack(fill="x")
        
        Label(sound_frame, text="提醒声音文件:").grid(row=0, column=0, sticky="w")
        self.sound_var = StringVar(value=self.config["alert_settings"]["sound_file"])
        Entry(sound_frame, textvariable=self.sound_var, width=30).grid(row=0, column=1, padx=5)
        Button(sound_frame, text="浏览", command=self.browse_sound).grid(row=0, column=2, padx=5)
        Button(sound_frame, text="测试", command=self.test_sound).grid(row=0, column=3, padx=5)
        
        # 提醒模式设置
        alert_frame = Frame(self.root, padx=10, pady=5)
        alert_frame.pack(fill="x")
        
        Label(alert_frame, text="提醒模式:").grid(row=0, column=0, sticky="w")
        
        self.alert_mode = StringVar(value=self.config["alert_settings"]["alert_mode"])
        
        self.popup_mode_btn = Button(
            alert_frame, 
            text="弹窗提醒", 
            command=lambda: self.set_alert_mode("popup"),
            width=10,
            bg="lightblue" if self.alert_mode.get() == "popup" else "SystemButtonFace"
        )
        self.popup_mode_btn.grid(row=0, column=1, padx=5)
        
        self.once_mode_btn = Button(
            alert_frame, 
            text="仅播放一次", 
            command=lambda: self.set_alert_mode("once"),
            width=10,
            bg="lightgreen" if self.alert_mode.get() == "once" else "SystemButtonFace"
        )
        self.once_mode_btn.grid(row=0, column=2, padx=5)
        
        # 检查间隔和自动开始设置
        settings_frame = Frame(self.root, padx=10, pady=5)
        settings_frame.pack(fill="x")
        
        Label(settings_frame, text="检查间隔(秒):").grid(row=0, column=0, sticky="w")
        self.interval_var = StringVar(value=str(self.config["alert_settings"]["check_interval"]))
        Entry(settings_frame, textvariable=self.interval_var, width=10).grid(row=0, column=1, padx=5)
        
        # 自动开始监控选项
        self.auto_start_var = BooleanVar(value=self.config["alert_settings"]["auto_start"])
        Checkbutton(settings_frame, text="程序启动时自动开始监控", variable=self.auto_start_var,
                   command=self.toggle_auto_start).grid(row=0, column=2, padx=20, sticky="w")
        
        # 控制按钮框架
        btn_frame = Frame(self.root, padx=10, pady=10)
        btn_frame.pack(fill="x")
        
        self.start_btn = Button(btn_frame, text="开始监控", command=self.start_monitor, 
                               bg="green", fg="white", font=("Arial", 10, "bold"), width=12, height=2)
        self.start_btn.pack(side="left", padx=5)
        
        self.stop_btn = Button(btn_frame, text="停止监控", command=self.stop_monitor, 
                              bg="red", fg="white", font=("Arial", 10, "bold"), width=12, height=2, state="disabled")
        self.stop_btn.pack(side="left", padx=5)
        
        Button(btn_frame, text="停止提醒", command=self.stop_alert, 
               bg="orange", fg="white", font=("Arial", 10, "bold"), width=12, height=2).pack(side="left", padx=5)
        
        Button(btn_frame, text="清空记录", command=self.clear_records, 
               bg="gray", fg="white", font=("Arial", 10, "bold"), width=12, height=2).pack(side="left", padx=5)
        
        # 保存设置按钮
        Button(btn_frame, text="保存设置", command=self.save_current_settings, 
               bg="blue", fg="white", font=("Arial", 10, "bold"), width=12, height=2).pack(side="left", padx=5)
        
        # 状态显示框架
        status_frame = Frame(self.root, padx=10, pady=10)
        status_frame.pack(fill="both", expand=True)
        
        Label(status_frame, text="运行状态:", font=("Arial", 10, "bold")).pack(anchor="w")
        
        # 创建带滚动条的文本框
        text_frame = Frame(status_frame)
        text_frame.pack(fill="both", expand=True, pady=(5,0))
        
        scrollbar = Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.status_text = Text(text_frame, yscrollcommand=scrollbar.set, wrap="word", height=10)
        self.status_text.pack(side="left", fill="both", expand=True)
        
        scrollbar.config(command=self.status_text.yview)
        
        # 初始状态
        self.update_status("程序已启动，配置已加载")
        self.update_status(f"已加载 {len(self.processed_emails)} 条邮件记录")
        self.update_status(f"当前提醒模式: {'弹窗提醒' if self.alert_mode.get() == 'popup' else '仅播放一次声音'}")
        
        if self.config["alert_settings"]["auto_start"]:
            self.update_status("自动开始监控已启用")
        else:
            self.update_status("请点击'开始监控'按钮或启用自动开始监控")
    
    def auto_start_monitoring(self):
        """自动开始监控"""
        # 检查必要的配置是否完整
        if not all([self.server_var.get(), self.port_var.get(), 
                   self.email_var.get(), self.password_var.get()]):
            self.update_status("自动开始失败: 邮箱设置不完整")
            return
        
        # 延迟2秒开始，确保界面完全加载
        self.root.after(2000, self.start_monitor)
    
    def toggle_auto_start(self):
        """切换自动开始设置"""
        auto_start = self.auto_start_var.get()
        if auto_start:
            self.update_status("已启用: 程序启动时自动开始监控")
        else:
            self.update_status("已禁用: 程序启动时自动开始监控")
    
    def set_alert_mode(self, mode):
        """设置提醒模式"""
        self.alert_mode.set(mode)
        
        # 更新按钮样式
        if mode == "popup":
            self.popup_mode_btn.config(bg="lightblue")
            self.once_mode_btn.config(bg="SystemButtonFace")
        else:
            self.popup_mode_btn.config(bg="SystemButtonFace")
            self.once_mode_btn.config(bg="lightgreen")
        
        self.update_status(f"提醒模式已切换: {'弹窗提醒' if mode == 'popup' else '仅播放一次声音'}")
    
    def browse_sound(self):
        """浏览选择声音文件"""
        filename = filedialog.askopenfilename(
            title="选择提醒声音文件",
            filetypes=[("WAV声音文件", "*.wav"), ("所有文件", "*.*")]
        )
        if filename:
            self.sound_var.set(filename)
    
    def test_sound(self):
        """测试声音文件"""
        sound_file = self.sound_var.get()
        if not os.path.exists(sound_file):
            messagebox.showerror("错误", "声音文件不存在！")
            return
        
        try:
            winsound.PlaySound(sound_file, winsound.SND_FILENAME)
            self.update_status("测试声音播放成功")
        except Exception as e:
            messagebox.showerror("错误", f"播放声音失败: {str(e)}")
    
    def update_status(self, msg):
        """更新状态显示"""
        current_time = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert("end", f"[{current_time}] {msg}\n")
        self.status_text.see("end")
        self.root.update_idletasks()
    
    def play_alert_once(self):
        """播放一次提醒声音"""
        sound_file = self.sound_var.get()
        if not os.path.exists(sound_file):
            self.update_status("错误: 声音文件不存在！")
            return
        
        try:
            winsound.PlaySound(sound_file, winsound.SND_FILENAME)
        except:
            pass
    
    def play_alert_loop(self):
        """循环播放提醒声音"""
        sound_file = self.sound_var.get()
        if not os.path.exists(sound_file):
            self.update_status("错误: 声音文件不存在！")
            return
        
        try:
            while self.is_alerting:
                winsound.PlaySound(sound_file, winsound.SND_FILENAME)
                for i in range(20):
                    if not self.is_alerting:
                        break
                    time.sleep(0.1)
        except:
            pass
    
    def stop_alert(self):
        """停止提醒"""
        self.is_alerting = False
        self.update_status("提醒已停止")
    
    def get_email_content(self, msg):
        """获取邮件内容"""
        content = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    try:
                        payload = part.get_payload(decode=True)
                        if payload:
                            content = payload.decode('utf-8', errors='ignore')
                            break
                    except:
                        pass
        else:
            content_type = msg.get_content_type()
            if content_type == "text/plain":
                try:
                    payload = msg.get_payload(decode=True)
                    if payload:
                        content = payload.decode('utf-8', errors='ignore')
                except:
                    pass
        return content
    
    def decode_header(self, header):
        """解码邮件头"""
        if not header:
            return ""
        try:
            decoded_parts = decode_header(header)
            decoded_str = ""
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding:
                        decoded_str += part.decode(encoding)
                    else:
                        decoded_str += part.decode('utf-8', errors='ignore')
                else:
                    decoded_str += part
            return decoded_str
        except:
            return str(header)
    
    def check_email(self):
        """检查新邮件"""
        check_interval = int(self.interval_var.get())
        
        while self.is_running:
            try:
                # 连接邮件服务器
                mail = imaplib.IMAP4_SSL(self.server_var.get(), int(self.port_var.get()))
                mail.login(self.email_var.get(), self.password_var.get())
                mail.select("INBOX")
                
                # 搜索未读邮件
                status, messages = mail.search(None, "UNSEEN")
                if status == "OK" and messages[0]:
                    email_ids = messages[0].split()
                    new_emails = []
                    
                    for email_id in email_ids:
                        status, msg_data = mail.fetch(email_id, '(RFC822)')
                        if status != "OK":
                            continue
                            
                        msg = email.message_from_bytes(msg_data[0][1])
                        subject = self.decode_header(msg['Subject'])
                        from_addr = self.decode_header(msg['From'])
                        date = msg['Date']
                        
                        email_data = {
                            'from': from_addr,
                            'subject': subject,
                            'date': date
                        }
                        email_hash = self.get_email_hash(email_data)
                        
                        if email_hash in self.processed_emails:
                            self.update_status(f"跳过已处理邮件: {subject}")
                            continue
                        
                        new_emails.append({
                            'id': email_id,
                            'hash': email_hash,
                            'subject': subject,
                            'from': from_addr,
                            'date': date
                        })
                    
                    if new_emails:
                        email_count = len(new_emails)
                        self.update_status(f"收到 {email_count} 封新邮件，开始提醒！")
                        
                        if self.alert_mode.get() == "once":
                            self.play_alert_once()
                            self.update_status("已播放一次提醒声音")
                            
                            for email_info in new_emails:
                                self.processed_emails.add(email_info['hash'])
                                self.update_status(f"标记邮件为已处理: {email_info['subject']}")
                            
                            self.save_processed_emails()
                            
                        else:
                            self.is_alerting = True
                            alert_thread = threading.Thread(target=self.play_alert_loop)
                            alert_thread.daemon = True
                            alert_thread.start()
                            
                            self.show_alert_dialog(email_count, new_emails)
                
                mail.close()
                mail.logout()
                
            except Exception as e:
                self.update_status(f"检查邮件错误: {str(e)}")
            
            for i in range(check_interval):
                if not self.is_running:
                    break
                time.sleep(1)
    
    def show_alert_dialog(self, count, new_emails):
        """显示提醒对话框"""
        dialog = Tk()
        dialog.title("新邮件提醒")
        dialog.geometry("400x200")
        dialog.attributes('-topmost', True)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        Label(dialog, text="📧 新邮件提醒", font=("Arial", 14, "bold"), fg="blue").pack(pady=10)
        Label(dialog, text=f"收到 {count} 封新邮件！", font=("Arial", 11)).pack(pady=5)
        Label(dialog, text="点击确认停止提醒", font=("Arial", 9)).pack(pady=5)
        
        def confirm():
            self.stop_alert()
            
            # 标记邮件为已处理
            for email_info in new_emails:
                self.processed_emails.add(email_info['hash'])
                self.update_status(f"标记邮件为已处理: {email_info['subject']}")
            
            self.save_processed_emails()
            dialog.destroy()
        
        Button(dialog, text="确认收到", command=confirm, 
               bg="green", fg="white", font=("Arial", 10, "bold"), width=10).pack(pady=10)
        
        self.alert_dialog = dialog
        dialog.mainloop()
    
    def clear_records(self):
        """清空已处理邮件记录"""
        if messagebox.askyesno("确认", "确定要清空所有已处理邮件记录吗？"):
            self.processed_emails.clear()
            self.save_processed_emails()
            self.update_status("已清空所有邮件记录")
    
    def save_current_settings(self):
        """保存当前设置到配置文件"""
        self.config["email_settings"] = {
            "server": self.server_var.get(),
            "port": self.port_var.get(),
            "email": self.email_var.get(),
            "password": self.password_var.get()
        }
        
        self.config["alert_settings"] = {
            "sound_file": self.sound_var.get(),
            "alert_mode": self.alert_mode.get(),
            "check_interval": int(self.interval_var.get()),
            "auto_start": self.auto_start_var.get()  # 保存自动开始设置
        }
        
        self.save_config()
        self.update_status("设置已保存")
    
    def start_monitor(self):
        """开始监控"""
        if not all([self.server_var.get(), self.port_var.get(), 
                   self.email_var.get(), self.password_var.get()]):
            messagebox.showerror("错误", "请填写完整的邮箱设置！")
            return
        
        # 保存当前设置
        self.save_current_settings()
        
        # 启动监控
        self.is_running = True
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        
        monitor_thread = threading.Thread(target=self.check_email)
        monitor_thread.daemon = True
        monitor_thread.start()
        
        self.update_status("邮件监控已启动，正在检查新邮件...")
        self.update_status(f"检查间隔: {self.interval_var.get()}秒")
        self.update_status(f"提醒模式: {'弹窗提醒' if self.alert_mode.get() == 'popup' else '仅播放一次声音'}")
    
    def stop_monitor(self):
        """停止监控"""
        self.is_running = False
        self.stop_alert()
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        
        if hasattr(self, 'alert_dialog'):
            try:
                self.alert_dialog.destroy()
            except:
                pass
        
        self.update_status("邮件监控已停止")
    
    def on_closing(self):
        """窗口关闭事件"""
        self.save_current_settings()
        
        if self.is_running:
            if messagebox.askokcancel("退出", "邮件监控正在运行，确定要退出吗？"):
                self.stop_monitor()
                self.root.destroy()
        else:
            self.root.destroy()
    
    def run(self):
        """运行程序"""
        self.root.mainloop()

def main():
    """主函数"""
    current_pid = os.getpid()
    current_script = os.path.basename(__file__)
    
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['pid'] != current_pid and proc.info['name'] in ['python.exe', 'pythonw.exe']:
                cmdline = proc.cmdline()
                if len(cmdline) > 1 and current_script in cmdline[1]:
                    messagebox.showwarning("警告", "程序已在运行中！")
                    return
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    
    # 创建默认声音文件
    if not os.path.exists("alert.wav"):
        try:
            import wave
            import struct
            import math
            
            sample_rate = 8000
            duration = 1.0
            frequency = 600
            
            with wave.open("alert.wav", "w") as wav_file:
                wav_file.setnchannels(1)
                wav_file.setsampwidth(2)
                wav_file.setframerate(sample_rate)
                
                frames = []
                for i in range(int(duration * sample_rate)):
                    value = int(32767.0 * math.sin(2 * math.pi * frequency * i / sample_rate))
                    frames.append(struct.pack('h', value))
                
                wav_file.writeframes(b''.join(frames))
                
            print("已创建默认提示音文件: alert.wav")
        except Exception as e:
            print(f"创建默认声音文件失败: {e}")
    
    app = EnhancedEmailAlert()
    app.run()

if __name__ == "__main__":
    main()