import subprocess
import os
import shutil
import time
import psutil
import tkinter as tk
from tkinter import messagebox

# 启动start程序
print(f"正在启动start11安装程序")
start_exe = os.path.abspath('../Tools/Start11.exe')
if start_exe and os.path.exists(start_exe):
        try:
            subprocess.run(start_exe, check=True)
            print("Start11安装成功")
        except subprocess.CalledProcessError:
            print(f"错误！ Start11 安装失败！")

# 间隔1秒后复制资源文件
time.sleep(1)
    
# 复制Start资源文件
print("准备进行start资源文件复制")

def clear_and_copy_files():
    response = messagebox.askquestion("清除和复制文件", "是否清除start自带资源文件并将咩咩主题文件复制至start资源文件夹")
    start_buttons_path = r"C:\Program Files (x86)\Stardock\Start11\StartButtons"
    menu_textures_path = r"C:\Program Files (x86)\Stardock\Start11\MenuTextures"
    taskbar_textures_path = r"C:\Program Files (x86)\Stardock\Start11\TaskbarTextures"

    if response == 'yes':
        # 清理Start11原资源文件
        clear_directory(start_buttons_path)
        clear_directory(menu_textures_path)
        clear_directory(taskbar_textures_path)

        res_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "Res")

        # 复制咩咩的Windows开始图标
        win_icon_folder = os.path.join(res_folder, "start", "win图标")
        copy_files_to_directory(win_icon_folder, start_buttons_path)

        # 复制咩咩的Windows开始菜单背景
        menu_background_folder = os.path.join(res_folder, "start", "开始背景")
        copy_files_to_directory(menu_background_folder, menu_textures_path)

        # 复制咩咩的Windows任务栏背景
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

    print("start资源文件复制完毕，请稍后参照教程手动调整设置")

# 间隔1秒后设置文件管理器背景
time.sleep(1)

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

    # 获取注册脚本路径
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # 定位注册文件路径
    cmd_file_path = os.path.join(script_directory, "..", "Tools", "Explorer Tool", "注册_Register.cmd")

    # 运行注册文件
    returncode, stdout, stderr = run_cmd(cmd_file_path)

    if returncode == 0:
        print("文件管理器背景设置成功，将在您下次打开文件管理器时显示")
        print(stdout)
    else:
        print("背景设置失败，原因如下")
        print("Error:")
        print(stderr)

# 间隔1秒后安装Rainmeter
time.sleep(1)

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

# 间隔1秒后复制主题文件
time.sleep(1)

# 设定Rainmeter主题文件源目录并根据当前用户设置目标文件夹
src_folder = os.path.abspath('../Res/Rainmeter')
username = os.getlogin()
documents_folder = os.path.join(os.path.expanduser('~'), 'Documents')
print(f"起始路径为: ", src_folder)
print(f"目标路径为：", documents_folder)
target_folder = os.path.join(os.path.expanduser('~'), 'Documents')

# 遍历目标文件夹
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

# 间隔1秒后重启Rainmeter程序
time.sleep(1)

# 获取Rainmeter的文件路径
def get_process_directory(process_name):
    process_directory = None
    process_found = False
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

# 设置进程名
process_name = "Rainmeter.exe"
process_directory = get_process_directory(process_name)
if process_directory:
    print("进程 '{}' 的目录为：".format(process_name), process_directory)
    print("准备重启Rainmeter进行主题刷新")

# 重启Rainmeter
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
                time.sleep(2)

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

# 间隔1秒后执行动态壁纸选择环节
time.sleep(1)

# 复制动态壁纸文件至C盘目录
def copy_background_files():
    source_background_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "Res", "background")
    target_background_path = "C:\\ThemeRes\\background"
    shutil.copytree(source_background_path, target_background_path)

    source_lock_screen_wallpaper_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "Res", "Lock screen wallpaper")
    target_lock_screen_wallpaper_path = "C:\\ThemeRes\\Lock screen wallpaper"
    shutil.copytree(source_lock_screen_wallpaper_path, target_lock_screen_wallpaper_path)

if __name__ == "__main__":
    copy_background_files()

# 执行动态壁纸选择
def on_main_window_click(option):
    main_window.destroy()

    if option == "yes":
        create_second_window()
    else:
        skip_step()

def run_lively():
    lively_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "Tools", "lively.exe")
    subprocess.run([lively_path])

def on_second_window_click(option):
    second_window.destroy()

    if option == "yes":
        open_background_folder()
    else:
        run_lively()
        wait_for_lively_completion()

def wait_for_lively_completion():
    messagebox.showinfo("添加壁纸", "请添加本目录中的壁纸至您的动态壁纸软件中")
    background_folder_path = os.path.join(os.path.dirname(os.path.abspath(__file__)))
    subprocess.Popen(["explorer", "C:\\ThemeRes\\background"])

def create_main_window():
    global main_window

    main_window = tk.Tk()
    main_window.title("选择是否需要动态壁纸")

    label = tk.Label(main_window, text="是否需要动态壁纸？")
    label.pack()

    button_yes = tk.Button(main_window, text="我需要", command=lambda: on_main_window_click("yes"))
    button_yes.pack()

    button_no = tk.Button(main_window, text="我不需要", command=lambda: on_main_window_click("no"))
    button_no.pack()

    main_window.mainloop()

def create_second_window():
    global second_window

    second_window = tk.Tk()
    second_window.title("是否拥有动态壁纸软件")

    label = tk.Label(second_window, text="是否拥有动态壁纸软件？")
    label.pack()

    button_yes = tk.Button(second_window, text="是，我拥有", command=lambda: on_second_window_click("yes"))
    button_yes.pack()

    button_no = tk.Button(second_window, text="不，我没有", command=lambda: on_second_window_click("no"))
    button_no.pack()

    second_window.mainloop()

def open_background_folder():
    background_folder_path = os.path.join(os.path.dirname(os.path.abspath(__file__)))
    subprocess.Popen(["explorer", "C:\\ThemeRes\\background"])

    messagebox.showinfo("添加壁纸", "请添加本目录中的壁纸至您的动态壁纸软件中")

def skip_step():
    messagebox.showinfo("跳过步骤", "已跳过添加动态壁纸步骤")

if __name__ == "__main__":
    create_main_window()
