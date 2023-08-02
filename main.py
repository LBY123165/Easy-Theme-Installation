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
    return wrappe

# 根据 Windows 版本选择要启动的程序
start_exe = os.path.abspath('../Tools/Start11.exe')
        
    # 打印获取到的启动程序路径，如果是None则打印相应的错误信息
print(f"正在启动start11安装程序")

    # 启动程序
if start_exe and os.path.exists(start_exe):
        try:
            subprocess.run(start_exe, check=True)
        except subprocess.CalledProcessError:
            print(f"错误！ {start_exe} 启动失败！")

# 进行start资源文件复制

print("准备进行start资源文件复制")

def clear_and_copy_files():
    response = messagebox.askquestion("清除和复制文件", "是否清除start自带资源文件并将咩咩主题文件复制至start资源文件夹")
    start_buttons_path = r"C:\Program Files (x86)\Stardock\Start11\StartButtons"
    menu_textures_path = r"C:\Program Files (x86)\Stardock\Start11\MenuTextures"
    taskbar_textures_path = r"C:\Program Files (x86)\Stardock\Start11\TaskbarTextures"

    if response == 'yes':
        # Clear files from start_buttons_path, menu_textures_path, and taskbar_textures_path
        clear_directory(start_buttons_path)
        clear_directory(menu_textures_path)
        clear_directory(taskbar_textures_path)

        res_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "Res")

        # Copy files from Res/start/win图标 to the respective directories
        win_icon_folder = os.path.join(res_folder, "start", "win图标")
        copy_files_to_directory(win_icon_folder, start_buttons_path)

        # Copy files from Res/start/开始背景 to the respective directories
        menu_background_folder = os.path.join(res_folder, "start", "开始背景")
        copy_files_to_directory(menu_background_folder, menu_textures_path)

        # Copy files from Res/start/任务栏 to the respective directories
        taskbar_folder = os.path.join(res_folder, "start", "任务栏")
        copy_files_to_directory(taskbar_folder, taskbar_textures_path)

        messagebox.showinfo("操作完成", "操作已完成！")
    else:
        res_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "Res")

        # Copy files from Res/start/win图标 to the respective directories
        win_icon_folder = os.path.join(res_folder, "start", "win图标")
        copy_files_to_directory(win_icon_folder, start_buttons_path)

        # Copy files from Res/start/开始背景 to the respective directories
        menu_background_folder = os.path.join(res_folder, "start", "开始背景")
        copy_files_to_directory(menu_background_folder, menu_textures_path)

        # Copy files from Res/start/任务栏 to the respective directories
        taskbar_folder = os.path.join(res_folder, "start", "任务栏")
        copy_files_to_directory(taskbar_folder, taskbar_textures_path)

        messagebox.showinfo("操作完成", "操作已完成！")

def clear_directory(directory_path):
    for filename in os.listdir(directory_path):
        if filename.endswith(".png"):
            file_path = os.path.join(directory_path, filename)
            os.remove(file_path)

def copy_files_to_directory(source_folder, destination_folder):
    for filename in os.listdir(source_folder):
        if filename.endswith(".png"):
            source_file_path = os.path.join(source_folder, filename)
            destination_file_path = os.path.join(destination_folder, filename)
            shutil.copyfile(source_file_path, destination_file_path)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    clear_and_copy_files()

    print("start资源文件复制完毕，请稍后手动调整设置")

# 设置资源管理器背景
def run_cmd(cmd):
    try:
        result = subprocess.run(cmd, shell=True, capture_output=False, text=True, check=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return result.returncode, result.stdout, result.stderr
    except subprocess.CalledProcessError as e:
        return e.returncode, e.stdout, e.stderr
    except Exception as e:
        return -1, "", str(e)

if __name__ == "__main__":
    # 获取当前脚本所在路径
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # 定位"注册_Register.cmd"文件路径
    cmd_file_path = os.path.join(script_directory, "..", "Tools", "Explorer Tool", "注册_Register.cmd")

    # 运行"注册_Register.cmd"文件
    returncode, stdout, stderr = run_cmd(cmd_file_path)

    if returncode == 0:
        print("文件管理器背景设置成功，将在您下次打开文件管理器时显示")
        print(stdout)
    else:
        print("背景设置失败，原因如下")
        print("Error:")
        print(stderr)

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
