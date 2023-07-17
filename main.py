import platform
import subprocess
import os
import winreg
import shutil
import psutil
import ctypes
import tkinter as tk
from tkinter import messagebox
import time

  # 函数装饰器，用于捕获函数执行中的异常
def handle_errors(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except:
            return None
    return wrapper


# 获取当前 Windows 内部版本号
@handle_errors
def get_windows_version():
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion') as key:
        version_number = winreg.QueryValueEx(key, 'CurrentBuildNumber')[0]
    return int(version_number)


# 获取当前 Windows 内部版本号
version_number = get_windows_version()

# 打印获取到的版本号，如果是None则打印相应的错误信息
if version_number is not None:
    if version_number <= 22000 >= 19000:
        print(f"当前系统版本为 Windows 10，将启动Start10安装程序")
    elif version_number >= 22000:
        print(f"当前系统版本为 Windows 11，将启动Start11安装程序")
    else:
        start_exe = None

# 根据 Windows 版本选择要启动的程序
if version_number is not None:
    if version_number <= 22000 >= 19000:
        start_exe = os.path.abspath('../Tools/Start10.exe')
    elif version_number >= 22000:
        start_exe = os.path.abspath('../Tools/Start11.exe')
    else:
        start_exe = None
        
    # 打印获取到的启动程序路径，如果是None则打印相应的错误信息
    if start_exe is not None:
        print(f"启动: {start_exe}")
    else:
        print(f"错误！无法找到 {start_exe}")

    # 启动程序
    if start_exe and os.path.exists(start_exe):
        try:
            subprocess.run(start_exe, check=True)
        except subprocess.CalledProcessError:
            print(f"错误！ {start_exe} 启动失败！")

    # 安装 Rainmeter
    rainmeter_install = os.path.abspath('../Tools/Rainmeter-install.exe')

    print(f"{start_exe}安装任务执行完毕，准备进行安装Rainmeter程序")
    
    if start_exe is not None:
        print(f"启动: {rainmeter_install}")
    else:
        print(f"错误！无法找到 {rainmeter_install}")

    # 启动程序
    if os.path.exists(rainmeter_install):
        try:
            subprocess.run(rainmeter_install, check=True)
        except subprocess.CalledProcessError:
            print(f"错误！ {rainmeter_install} 运行失败！")
    else:
        print(f"错误！无法找到 {rainmeter_install}")

    # Rainmeter 后续安装
    print(f"Rainmeter安装任务执行成功，准备复制主题文件")

# 指定源文件夹路径
src_folder = os.path.abspath('../Res/Rainmeter')
  
# 设置用户目录
# 获取当前登录用户的用户名
username = os.getlogin()
documents_folder = os.path.join(os.path.expanduser('~'), 'Documents')
print(f"起始路径为: ", src_folder)
print(f"目标路径为：", documents_folder)
        
# 目标文件夹路径
target_folder = os.path.join(os.path.expanduser('~'), 'Documents')

    # 遍历源文件夹中的文件和文件夹
for item in os.listdir(src_folder):
        source_item = os.path.join(src_folder, item)
        target_item = os.path.join(target_folder, item)
    
        # 检查目标文件是否存在
        if os.path.exists(target_item):
            # 如果是文件夹，则递归替换文件夹
            if os.path.isdir(source_item):
                shutil.rmtree(target_item)
                shutil.copytree(source_item, target_item)
                print(f"文件夹 {item} 替换成功！")
            # 如果是文件，则替换文件
            elif os.path.isfile(source_item):
                shutil.copy(source_item, target_item)
                print(f"文件 {item} 替换成功！")
        else:
            # 如果目标文件不存在，则直接复制文件或文件夹
            shutil.copytree(source_item, target_item) if os.path.isdir(source_item) else shutil.copy(source_item, target_item)
            print(f"文件夹 {item} 复制成功！")


def get_process_directory(process_name):
    process_directory = None
    process_found = False

    # 遍历所有正在运行的进程
    for process in psutil.process_iter(['pid', 'name']):
        try:
            # 检查进程名是否匹配
            if process.info['name'].lower() == process_name.lower():
                process_found = True
                # 获取进程的可执行文件路径
                process_executable = process.exe()

                # 提取目录部分，去除文件名，只保留目录路径
                process_directory = os.path.dirname(process_executable)

                break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # 忽略无法访问的进程
            pass

    if not process_found:
        print("未找到进程 '{}'".format(process_name))

    return process_directory

# 指定进程名
process_name = "Rainmeter.exe"
process_directory = get_process_directory(process_name)
if process_directory:
    print("进程 '{}' 的目录为：".format(process_name), process_directory)
    print("准备重启Rainmeter进行主题刷新")



def restart_process(process_name):
    # 遍历所有正在运行的进程
    for process in psutil.process_iter(['pid', 'name']):
        try:
            # 检查进程名是否匹配
            if process.info['name'].lower() == process_name.lower():
                process_id = process.info['pid']

                # 使用 taskkill 命令结束进程
                subprocess.run(['taskkill', '/F', '/PID', str(process_id)], stdout=subprocess.PIPE, text=True)

                # 这里可以加上一些等待时间，确保进程完全终止
                # time.sleep(2)

                # 启动进程，例如使用 os.system() 调用
                os.system(process_name)
                break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # 忽略无法访问的进程
            pass

# 指定进程名
src_folder = process_directory
process_name = "start Rainmeter.exe"
restart_process(process_name)

print("Rainmeter重启完毕，请稍后手动开启插件")

#动态壁纸选择
def install_dynamic_wallpaper():
    # 这里是选项一的操作：运行上级目录中的 Tools 文件夹中的 lively.exe
    tools_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Tools"))
    lively_exe_path = os.path.join(tools_dir, "lively.exe")

    if os.path.exists(lively_exe_path):
        print("正在进行安装...")
        # 弹出新的窗口显示即将开始安装信息
        info_window = tk.Tk()
        info_window.title("安装提示")
        info_label = tk.Label(info_window, text="即将开始安装lively动态壁纸软件")
        info_label.pack(padx=20, pady=10)

        # 停留 3 秒
        info_window.after(3000, lambda: run_lively_and_exit(lively_exe_path, info_window))
    else:
        messagebox.showerror("错误", "未找到 lively.exe，请确保文件存在于 Tools 文件夹中。")

def run_lively_and_exit(lively_exe_path, window_to_close):
    window_to_close.destroy()  # 关闭信息窗口
    subprocess.run(lively_exe_path)
    print("安装过程结束")

def already_have_wallpaper():
    # 这里是选项二的操作
    # 添加选项二的具体操作，或者留空表示不进行任何操作
    print("安装过程结束")

def not_needed():
    # 这里是选项三的操作
    # 添加选项三的具体操作，或者留空表示不进行任何操作
    print("安装过程结束")

def show_choice_window():
    # 创建主窗口
    root = tk.Tk()
    root.title("选择窗口")

    # 定义选择窗口的回调函数
    def on_yes():
        install_dynamic_wallpaper()
        root.destroy()  # 销毁选择窗口

    def on_no_have():
        already_have_wallpaper()
        root.destroy()  # 销毁选择窗口

    def on_no_need():
        not_needed()
        root.destroy()  # 销毁选择窗口

    # 创建选择窗口的内容
    label = tk.Label(root, text="是否安装动态壁纸软件")
    label.pack(padx=20, pady=10)

    yes_button = tk.Button(root, text="是", command=on_yes)
    yes_button.pack(pady=5)

    no_have_button = tk.Button(root, text="不，我已经有了", command=on_no_have)
    no_have_button.pack(pady=5)

    no_need_button = tk.Button(root, text="不，我不需要", command=on_no_need)
    no_need_button.pack(pady=5)

    # 进入主循环
    root.mainloop()

# 调用函数显示选择窗口
show_choice_window()
