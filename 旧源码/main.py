import tkinter as tk
from tkinter import ttk, messagebox
import wmi
import json
import os
from tkinter import filedialog
from tkinter.font import Font
import ctypes
import sys

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

class NetworkConfigTool:
    def __init__(self, root):
        if not is_admin():
            messagebox.showerror("错误", "此程序需要管理员权限才能修改网络设置。\n程序将以管理员身份重新启动。")
            run_as_admin()
            root.destroy()
            return
            
        self.root = root
        self.root.title("网络适配器配置工具")
        self.root.geometry("800x700")  # 增加窗口高度
        self.root.minsize(800, 700)    # 设置最小窗口大小
        
        # 设置主题颜色
        self.bg_color = "#f0f0f0"
        self.accent_color = "#2196F3"  # Material Design Blue
        self.text_color = "#333333"
        
        # 设置字体
        self.title_font = Font(family="Microsoft YaHei UI", size=12, weight="bold")
        self.normal_font = Font(family="Microsoft YaHei UI", size=10)
        
        self.root.configure(bg=self.bg_color)
        
        # 创建样式
        self.create_styles()
        
        self.wmi = wmi.WMI()
        self.create_widgets()
        self.load_adapters()
        
        # 配置存储路径
        self.config_dir = "network_configs"
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)

    def create_styles(self):
        # 创建自定义样式
        style = ttk.Style()
        style.configure("Title.TLabel", 
                       font=self.title_font, 
                       background=self.bg_color, 
                       foreground=self.text_color)
        
        style.configure("Custom.TFrame", 
                       background=self.bg_color)
        
        style.configure("Custom.TLabelframe", 
                       background=self.bg_color)
        
        style.configure("Custom.TLabelframe.Label", 
                       font=self.title_font,
                       background=self.bg_color, 
                       foreground=self.text_color)
        
        # 按钮样式
        style.configure("Accent.TButton",
                       font=self.normal_font,
                       padding=10)
        
        # 输入框样式
        style.configure("Custom.TEntry",
                       font=self.normal_font,
                       padding=5)

    def create_widgets(self):
        # 主容器
        main_container = ttk.Frame(self.root, style="Custom.TFrame")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_frame = ttk.Frame(main_container, style="Custom.TFrame")
        title_frame.pack(fill="x", pady=(0, 20))
        
        title_label = ttk.Label(title_frame, 
                              text="网络适配器配置工具", 
                              style="Title.TLabel")
        title_label.pack()
        
        # 状态显示区域
        self.status_frame = ttk.LabelFrame(main_container,
                                         text="适配器状态",
                                         style="Custom.TLabelframe",
                                         padding="10")
        self.status_frame.pack(fill="x", pady=(0, 10))
        
        # 状态标签
        self.status_label = ttk.Label(self.status_frame,
                                    text="",
                                    font=self.normal_font,
                                    background=self.bg_color,
                                    wraplength=700)  # 设置文本自动换行
        self.status_label.pack(fill="x", pady=5)
        
        # 适配器选择
        adapter_frame = ttk.LabelFrame(main_container, 
                                     text="网络适配器", 
                                     style="Custom.TLabelframe",
                                     padding="10")
        adapter_frame.pack(fill="x", pady=(0, 10))
        
        self.adapter_combo = ttk.Combobox(adapter_frame, 
                                        width=50,
                                        font=self.normal_font)
        self.adapter_combo.pack(fill="x", pady=5)
        self.adapter_combo.bind('<<ComboboxSelected>>', self.on_adapter_selected)
        
        # IP设置
        ip_frame = ttk.LabelFrame(main_container, 
                                text="网络设置", 
                                style="Custom.TLabelframe",
                                padding="10")
        ip_frame.pack(fill="x", pady=(0, 10))
        
        # 创建网格布局
        for i in range(5):
            ip_frame.grid_columnconfigure(1, weight=1)
        
        # IP地址
        ttk.Label(ip_frame, 
                 text="IP地址:", 
                 font=self.normal_font,
                 background=self.bg_color).grid(row=0, column=0, sticky="w", pady=8)
        self.ip_entry = ttk.Entry(ip_frame, font=self.normal_font)
        self.ip_entry.grid(row=0, column=1, sticky="ew", padx=(10, 0))
        
        # 子网掩码
        ttk.Label(ip_frame, 
                 text="子网掩码:", 
                 font=self.normal_font,
                 background=self.bg_color).grid(row=1, column=0, sticky="w", pady=8)
        self.subnet_entry = ttk.Entry(ip_frame, font=self.normal_font)
        self.subnet_entry.grid(row=1, column=1, sticky="ew", padx=(10, 0))
        
        # 默认网关
        ttk.Label(ip_frame, 
                 text="默认网关:", 
                 font=self.normal_font,
                 background=self.bg_color).grid(row=2, column=0, sticky="w", pady=8)
        self.gateway_entry = ttk.Entry(ip_frame, font=self.normal_font)
        self.gateway_entry.grid(row=2, column=1, sticky="ew", padx=(10, 0))
        
        # DNS服务器
        ttk.Label(ip_frame, 
                 text="首选DNS:", 
                 font=self.normal_font,
                 background=self.bg_color).grid(row=3, column=0, sticky="w", pady=8)
        self.dns1_entry = ttk.Entry(ip_frame, font=self.normal_font)
        self.dns1_entry.grid(row=3, column=1, sticky="ew", padx=(10, 0))
        
        ttk.Label(ip_frame, 
                 text="备用DNS:", 
                 font=self.normal_font,
                 background=self.bg_color).grid(row=4, column=0, sticky="w", pady=8)
        self.dns2_entry = ttk.Entry(ip_frame, font=self.normal_font)
        self.dns2_entry.grid(row=4, column=1, sticky="ew", padx=(10, 0))
        
        # 按钮区域
        button_frame = ttk.Frame(main_container, style="Custom.TFrame")
        button_frame.pack(fill="x", pady=(20, 0), padx=5)
        
        # 创建按钮容器，使用网格布局
        button_grid = ttk.Frame(button_frame, style="Custom.TFrame")
        button_grid.pack(fill="x")
        button_grid.grid_columnconfigure((0,1,2,3), weight=1)  # 平均分配空间
        
        # 创建按钮并使用网格布局
        apply_btn = ttk.Button(button_grid, 
                             text="应用设置", 
                             command=self.apply_settings,
                             style="Accent.TButton")
        apply_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        dhcp_btn = ttk.Button(button_grid,
                            text="启用DHCP",
                            command=self.enable_dhcp,
                            style="Accent.TButton")
        dhcp_btn.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        save_btn = ttk.Button(button_grid, 
                            text="保存配置", 
                            command=self.save_config,
                            style="Accent.TButton")
        save_btn.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        
        load_btn = ttk.Button(button_grid, 
                            text="加载配置", 
                            command=self.load_config,
                            style="Accent.TButton")
        load_btn.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

    def load_adapters(self):
        """加载系统中的网络适配器"""
        self.adapters = self.wmi.Win32_NetworkAdapter(PhysicalAdapter=True)
        adapter_names = [adapter.NetConnectionID for adapter in self.adapters if hasattr(adapter, 'NetConnectionID')]
        self.adapter_combo['values'] = adapter_names
        if adapter_names:
            self.adapter_combo.set(adapter_names[0])
            self.on_adapter_selected(None)

    def on_adapter_selected(self, event):
        """当选择适配器时更新显示的配置"""
        selected = self.adapter_combo.get()
        for adapter in self.adapters:
            if hasattr(adapter, 'NetConnectionID') and adapter.NetConnectionID == selected:
                # 获取IP配置
                config = self.wmi.Win32_NetworkAdapterConfiguration(Index=adapter.Index)[0]
                
                # 更新状态显示
                status_text = f"适配器: {selected}\n"
                status_text += f"状态: {'已启用' if adapter.NetEnabled else '已禁用'}\n"
                status_text += f"IP获取方式: {'DHCP' if config.DHCPEnabled else '静态IP'}\n"
                status_text += f"MAC地址: {adapter.MACAddress if hasattr(adapter, 'MACAddress') else '未知'}\n"
                
                # 修复速度显示的问题
                speed = "未知"
                if hasattr(adapter, 'Speed') and adapter.Speed:
                    try:
                        speed_value = float(adapter.Speed) / 1000000
                        speed = f"{speed_value:.1f} Mbps"
                    except (ValueError, TypeError):
                        speed = "未知"
                status_text += f"连接速度: {speed}"
                
                # 如果是静态IP，显示更多信息
                if not config.DHCPEnabled and config.IPAddress:
                    status_text += f"\nIP地址: {config.IPAddress[0]}"
                    status_text += f"\n子网掩码: {config.IPSubnet[0]}"
                    if config.DefaultIPGateway:
                        status_text += f"\n默认网关: {config.DefaultIPGateway[0]}"
                    if config.DNSServerSearchOrder:
                        status_text += f"\nDNS服务器: {', '.join(config.DNSServerSearchOrder)}"
                
                self.status_label.config(text=status_text)
                
                if config.IPAddress:
                    self.ip_entry.delete(0, tk.END)
                    self.ip_entry.insert(0, config.IPAddress[0])
                    self.subnet_entry.delete(0, tk.END)
                    self.subnet_entry.insert(0, config.IPSubnet[0])
                if config.DefaultIPGateway:
                    self.gateway_entry.delete(0, tk.END)
                    self.gateway_entry.insert(0, config.DefaultIPGateway[0])
                if config.DNSServerSearchOrder:
                    self.dns1_entry.delete(0, tk.END)
                    self.dns1_entry.insert(0, config.DNSServerSearchOrder[0])
                    if len(config.DNSServerSearchOrder) > 1:
                        self.dns2_entry.delete(0, tk.END)
                        self.dns2_entry.insert(0, config.DNSServerSearchOrder[1])

    def apply_settings(self):
        """应用网络设置"""
        if not is_admin():
            messagebox.showerror("错误", "需要管理员权限才能修改网络设置")
            return
            
        try:
            selected = self.adapter_combo.get()
            for adapter in self.adapters:
                if hasattr(adapter, 'NetConnectionID') and adapter.NetConnectionID == selected:
                    config = self.wmi.Win32_NetworkAdapterConfiguration(Index=adapter.Index)[0]
                    
                    # 获取当前配置
                    current_ip = config.IPAddress[0] if config.IPAddress else ""
                    current_subnet = config.IPSubnet[0] if config.IPSubnet else ""
                    current_gateway = config.DefaultIPGateway[0] if config.DefaultIPGateway else ""
                    current_dns = config.DNSServerSearchOrder if config.DNSServerSearchOrder else []
                    
                    # 获取用户输入
                    new_ip = self.ip_entry.get().strip()
                    new_subnet = self.subnet_entry.get().strip()
                    new_gateway = self.gateway_entry.get().strip()
                    new_dns1 = self.dns1_entry.get().strip()
                    new_dns2 = self.dns2_entry.get().strip()
                    new_dns = [dns for dns in [new_dns1, new_dns2] if dns]
                    
                    # 检查是否有任何设置被修改
                    settings_changed = (
                        new_ip != current_ip or
                        new_subnet != current_subnet or
                        new_gateway != current_gateway or
                        new_dns != current_dns
                    )
                    
                    if not settings_changed:
                        messagebox.showinfo("提示", "设置未发生改变")
                        return
                    
                    # 如果有任何设置被修改，就切换到静态IP模式
                    # 设置IP和子网掩码
                    if not new_ip or not new_subnet:
                        messagebox.showerror("错误", "IP地址和子网掩码不能为空")
                        return
                    
                    # 应用静态IP设置
                    result = config.EnableStatic([new_ip], [new_subnet])
                    if result[0] != 0:
                        messagebox.showerror("错误", f"设置IP地址失败，错误代码: {result[0]}")
                        return
                    
                    # 等待设置生效
                    import time
                    time.sleep(2)  # 增加等待时间
                    
                    # 设置网关
                    if new_gateway:
                        gateway_result = config.SetGateways([new_gateway])
                        if gateway_result[0] != 0:
                            messagebox.showwarning("警告", f"设置网关失败，错误代码: {gateway_result[0]}")
                    
                    time.sleep(1)
                    
                    # 设置DNS
                    if new_dns:
                        dns_result = config.SetDNSServerSearchOrder(new_dns)
                        if dns_result[0] != 0:
                            messagebox.showwarning("警告", f"设置DNS失败，错误代码: {dns_result[0]}")
                    
                    # 再次等待以确保所有设置都已应用
                    time.sleep(2)
                    
                    # 重新获取配置以更新状态
                    config = self.wmi.Win32_NetworkAdapterConfiguration(Index=adapter.Index)[0]
                    
                    # 验证设置是否成功
                    if config.IPAddress and config.IPAddress[0] == new_ip:
                        messagebox.showinfo("成功", "网络设置已更新为静态IP模式")
                    else:
                        messagebox.showwarning("警告", "IP设置可能未完全生效，请检查网络状态")
                    
                    # 刷新状态显示
                    self.on_adapter_selected(None)
                    break
        except Exception as e:
            messagebox.showerror("错误", f"设置失败: {str(e)}")

    def save_config(self):
        """保存当前配置到文件"""
        config = {
            'adapter': self.adapter_combo.get(),
            'ip': self.ip_entry.get(),
            'subnet': self.subnet_entry.get(),
            'gateway': self.gateway_entry.get(),
            'dns1': self.dns1_entry.get(),
            'dns2': self.dns2_entry.get()
        }
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=self.config_dir
        )
        
        if filename:
            with open(filename, 'w') as f:
                json.dump(config, f, indent=4)
            messagebox.showinfo("成功", "配置已保存")

    def load_config(self):
        """从文件加载配置"""
        filename = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=self.config_dir
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    config = json.load(f)
                
                self.adapter_combo.set(config['adapter'])
                self.ip_entry.delete(0, tk.END)
                self.ip_entry.insert(0, config['ip'])
                self.subnet_entry.delete(0, tk.END)
                self.subnet_entry.insert(0, config['subnet'])
                self.gateway_entry.delete(0, tk.END)
                self.gateway_entry.insert(0, config['gateway'])
                self.dns1_entry.delete(0, tk.END)
                self.dns1_entry.insert(0, config['dns1'])
                self.dns2_entry.delete(0, tk.END)
                self.dns2_entry.insert(0, config['dns2'])
                
                messagebox.showinfo("成功", "配置已加载")
            except Exception as e:
                messagebox.showerror("错误", f"加载配置失败: {str(e)}")

    def enable_dhcp(self):
        """启用DHCP"""
        if not is_admin():
            messagebox.showerror("错误", "需要管理员权限才能修改网络设置")
            return
        try:
            selected = self.adapter_combo.get()
            for adapter in self.adapters:
                if hasattr(adapter, 'NetConnectionID') and adapter.NetConnectionID == selected:
                    config = self.wmi.Win32_NetworkAdapterConfiguration(Index=adapter.Index)[0]
                    
                    # 启用DHCP
                    result = config.EnableDHCP()
                    
                    # 等待DHCP设置生效
                    import time
                    time.sleep(1)
                    
                    # 启用DNS DHCP
                    config.SetDNSServerSearchOrder()
                    
                    # 再次等待以确保设置生效
                    time.sleep(1)
                    
                    # 重新获取配置以确认DHCP状态
                    config = self.wmi.Win32_NetworkAdapterConfiguration(Index=adapter.Index)[0]
                    
                    if config.DHCPEnabled:
                        # 清空输入框
                        self.ip_entry.delete(0, tk.END)
                        self.subnet_entry.delete(0, tk.END)
                        self.gateway_entry.delete(0, tk.END)
                        self.dns1_entry.delete(0, tk.END)
                        self.dns2_entry.delete(0, tk.END)
                        
                        messagebox.showinfo("成功", "已成功启用DHCP")
                    else:
                        messagebox.showwarning("警告", "DHCP可能未完全启用，请检查网络状态")
                    
                    # 刷新显示的配置和状态
                    self.on_adapter_selected(None)
                    break
        except Exception as e:
            messagebox.showerror("错误", f"启用DHCP失败: {str(e)}")

if __name__ == '__main__':
    root = tk.Tk()
    app = NetworkConfigTool(root)
    root.mainloop()
