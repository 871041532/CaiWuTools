import os

def run():
    URL = "https://codeload.github.com/871041532/CaiWuTools/zip/master"
    ZIP_FILE = "code.zip"
    CODE_DIR = "code"

    # 基础模块
    print("检测基础模块...")
    module_names = [
        ("requests", "requests"),
        ("zipfile", "zipfile"),
        ("shutil", "shutil"),
    ]
    for module_name in module_names:
        try:
            __import__(module_name[0])
        except:
            print("  缺少module:" + module_name[0] + ", 开始安装...")
            os.system("pip install " + module_name[1])
            print("  安装 " + module_name[0] + " 完毕")
    import requests
    import zipfile
    import shutil
    print("基础模块ok")

    #下载
    print("1.开始下载...")
    r = requests.get(URL)
    with open(ZIP_FILE, 'wb') as f:
        f.write(r.content)
    print("下载完毕")

    # 解压
    print("2.开始解压...")
    if os.path.exists(CODE_DIR):
        shutil.rmtree(CODE_DIR)
    azip = zipfile.ZipFile(ZIP_FILE)
    azip.extractall(CODE_DIR)
    azip.close()
    print("解压完毕")

    # 文件转移
    print("3.文件开始转移...")
    master_dir = "./" + CODE_DIR +"/CaiWuTools-master/"
    files = os.listdir(master_dir)
    for file_name in files:
        shutil.move(master_dir + file_name, "./" + file_name)
    print("文件转移完毕")

    # 删除文件
    print("4.删除多余文件...")
    shutil.rmtree(CODE_DIR)
    os.remove(ZIP_FILE)
    print("多余文件删除完毕")

    # 安装缺失包
    print("5.检查需要的模块...")
    from Globals import Globals
    module_names = Globals.module_names
    for module_name in module_names:
        try:
            __import__(module_name[0])
        except:
            print("  缺少module:" + module_name[0] + ", 开始安装...")
            os.system("pip install " + module_name[1])
            print("  安装 " + module_name[0] + " 完毕")
    print("模块检测完毕")

    print("\n更新完毕, 按回车结束。")
    a = input()
run()