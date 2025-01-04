import sys
import os
import ctypes
import subprocess

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin(script_path):
    """以管理员权限运行指定的脚本"""
    if is_admin():
        # 如果已经是管理员权限，直接运行网络配置程序
        subprocess.run([sys.executable, script_path], check=True)
    else:
        # 请求管理员权限
        ctypes.windll.shell32.ShellExecuteW(
            None,           # 父窗口句柄
            "runas",       # 操作
            sys.executable,  # 程序
            script_path,    # 参数
            None,          # 默认目录
            1              # 显示窗口
        )

if __name__ == '__main__':
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 构建网络配置程序的路径
    network_config_path = os.path.join(current_dir, 'network_config.py')
    
    # 运行网络配置程序（使用管理员权限）
    run_as_admin(network_config_path)
