import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

def get_network_interfaces():
    """获取系统网络适配器列表"""
    try:
        result = subprocess.run(
            ['netsh', 'interface', 'show', 'interface'], 
            capture_output=True,
            encoding='gbk'  # 使用GBK编码处理中文Windows输出
        )
        
        if result.returncode != 0:
            print(f"获取网络适配器失败，错误代码: {result.returncode}")
            print(result.stderr)
            return []
            
        interfaces = []
        for line in result.stdout.splitlines()[3:]:  # 跳过表头
            if line.strip():
                parts = line.split()
                if len(parts) > 3:
                    interface_name = ' '.join(parts[3:])
                    interfaces.append(interface_name)
        return interfaces
        
    except Exception as e:
        print(f"获取网络适配器时发生错误: {str(e)}")
        return []

# 网络配置参数
STATIC_IP = "192.168.100.159"
SUBNET_MASK = "255.255.255.0"
GATEWAY = "192.168.100.10"
DNS1 = "192.168.100.1"
DNS2 = ""

def run_as_admin():
    """以管理员权限运行"""
    if sys.platform == 'win32':
        try:
            # 检查是否已经是管理员
            if os.getuid() == 0:
                return True
        except AttributeError:
            pass
        
        # 请求管理员权限
        ctypes = __import__('ctypes')
        shell32 = ctypes.windll.shell32
        if shell32.IsUserAnAdmin():
            return True
            
        # 重新以管理员身份运行
        shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        sys.exit()

def set_static_ip(interface_name):
    """设置静态IP"""
    print(f"正在为 {interface_name} 配置静态IP...")
    print(f"IP地址 = {STATIC_IP}")
    print(f"子网掩码 = {SUBNET_MASK}")
    print(f"网关 = {GATEWAY}")
    
    # 设置IP地址
    subprocess.run([
        'netsh', 'interface', 'ipv4', 'set', 'address',
        interface_name, 'static', STATIC_IP, SUBNET_MASK, GATEWAY
    ])
    
    # 设置DNS
    if DNS1:
        print(f"首选DNS = {DNS1}")
        subprocess.run([
            'netsh', 'interface', 'ipv4', 'set', 'dns',
            interface_name, 'static', DNS1
        ])
    else:
        print("DNS1为空")
        
    if DNS2:
        print(f"备用DNS = {DNS2}")
        subprocess.run([
            'netsh', 'interface', 'ipv4', 'add', 'dns',
            interface_name, DNS2
        ])
    else:
        print("DNS2为空")
        
    print(f"**********已设置为静态IP：{STATIC_IP}***********")

def set_dynamic_ip(interface_name):
    """设置动态IP"""
    print(f"正在为 {interface_name} 配置动态IP...")
    print("正在从DHCP自动获取IP地址...")
    subprocess.run([
        'netsh', 'interface', 'ip', 'set', 'address',
        interface_name, 'dhcp'
    ])
    
    print("正在从DHCP自动获取DNS地址...")
    subprocess.run([
        'netsh', 'interface', 'ip', 'set', 'dns',
        interface_name, 'dhcp'
    ])
    
    print("**********已设置为动态IP地址***********")

class NetworkConfigApp:
    def __init__(self, root):
        self.root = root
        self.root.title("网络配置工具")
        self.root.geometry("400x350")
        
        # 获取网络适配器列表
        self.interfaces = get_network_interfaces()
        
        # 网络适配器选择
        interface_frame = tk.LabelFrame(root, text="网络适配器", padx=10, pady=10)
        interface_frame.pack(fill="x", padx=10, pady=5)
        
        self.interface_var = tk.StringVar(value=self.interfaces[0] if self.interfaces else "")
        
        tk.Label(interface_frame, text="选择网络适配器:").grid(row=0, column=0)
        self.interface_menu = tk.OptionMenu(interface_frame, self.interface_var, *self.interfaces)
        self.interface_menu.grid(row=0, column=1, sticky="ew")
        
        # 静态IP配置框架
        static_frame = tk.LabelFrame(root, text="静态IP配置", padx=10, pady=10)
        static_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(static_frame, text="IP地址:").grid(row=0, column=0)
        self.ip_entry = tk.Entry(static_frame)
        self.ip_entry.insert(0, STATIC_IP)
        self.ip_entry.grid(row=0, column=1)
        
        tk.Label(static_frame, text="子网掩码:").grid(row=1, column=0)
        self.mask_entry = tk.Entry(static_frame)
        self.mask_entry.insert(0, SUBNET_MASK)
        self.mask_entry.grid(row=1, column=1)
        
        tk.Label(static_frame, text="网关:").grid(row=2, column=0)
        self.gateway_entry = tk.Entry(static_frame)
        self.gateway_entry.insert(0, GATEWAY)
        self.gateway_entry.grid(row=2, column=1)
        
        tk.Label(static_frame, text="DNS1:").grid(row=3, column=0)
        self.dns1_entry = tk.Entry(static_frame)
        self.dns1_entry.insert(0, DNS1)
        self.dns1_entry.grid(row=3, column=1)
        
        tk.Label(static_frame, text="DNS2:").grid(row=4, column=0)
        self.dns2_entry = tk.Entry(static_frame)
        self.dns2_entry.grid(row=4, column=1)
        
        tk.Button(static_frame, text="应用静态IP", command=self.apply_static_ip).grid(row=5, columnspan=2)
        
        # 动态IP配置
        dynamic_frame = tk.LabelFrame(root, text="动态IP配置", padx=10, pady=10)
        dynamic_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Button(dynamic_frame, text="应用动态IP", command=self.apply_dynamic_ip).pack()
        
        # 退出按钮
        tk.Button(root, text="退出", command=root.quit).pack(side="bottom", pady=10)

    def apply_static_ip(self):
        """应用静态IP配置"""
        global STATIC_IP, SUBNET_MASK, GATEWAY, DNS1, DNS2
        STATIC_IP = self.ip_entry.get()
        SUBNET_MASK = self.mask_entry.get()
        GATEWAY = self.gateway_entry.get()
        DNS1 = self.dns1_entry.get()
        DNS2 = self.dns2_entry.get()
        
        interface_name = self.interface_var.get()
        set_static_ip(interface_name)
        messagebox.showinfo("成功", f"已为 {interface_name} 应用静态IP配置")

    def apply_dynamic_ip(self):
        """应用动态IP配置"""
        interface_name = self.interface_var.get()
        set_dynamic_ip(interface_name)
        messagebox.showinfo("成功", f"已为 {interface_name} 应用动态IP配置")

if __name__ == "__main__":
    run_as_admin()
    root = tk.Tk()
    app = NetworkConfigApp(root)
    root.mainloop()
