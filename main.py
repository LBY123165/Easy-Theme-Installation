import platform
import subprocess
import os
import winreg
import shutil
import psutil
import ctypes



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

    print(f"{start_exe}安装成功，准备进行安装Rainmeter程序")
    
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
    print(f"Rainmeter安装成功，准备复制主题文件")

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

    process_name = "Rainmeter.exe"

    def get_process_directory(process_name):
        process_directory = None
        process_found = False

        # 遍历所有正在运行的进程
    for process in psutil.process_iter(['pid', 'name']):
        try:
            # 检查进程名是否匹配
            if process.info['name'].lower() == process_name.lower():
                process_found = True
                process_id = process.info['pid']

                # 使用 Windows API 获取进程句柄
                hProcess = ctypes.windll.kernel32.OpenProcess(0x1000, False, process_id)

                # 创建缓冲区来存储进程目录
                buf = ctypes.create_unicode_buffer(512)

                # 调用 Windows API 函数获取进程目录
                ctypes.windll.kernel32.GetModuleFileNameExW(hProcess, None, buf, ctypes.sizeof(buf))

                # 从缓冲区中获取进程目录
                process_directory = buf.value

                # 去除文件名部分，只保留目录路径
                process_directory = process_directory.rsplit('\\', 1)[0]

                # 关闭进程句柄
                ctypes.windll.kernel32.CloseHandle(hProcess)

                break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # 忽略无法访问的进程
            pass

    if not process_found:
        print("未找到进程 '{}'".format(process_name))


    # 指定进程名
    process_directory = get_process_directory(process_name)
    if process_directory:
        print("进程 '{}' 的目录：".format(process_name), process_directory)