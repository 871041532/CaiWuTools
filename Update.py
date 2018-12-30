import os
import requests
import zipfile
import shutil

def run():
    URL = "https://codeload.github.com/871041532/CaiWuTools/zip/master"
    ZIP_FILE = "code.zip"
    CODE_DIR = "code"

    #下载
    r = requests.get(URL)
    with open(ZIP_FILE, 'wb') as f:
        f.write(r.content)

    # 解压
    if os.path.exists(CODE_DIR):
        shutil.rmtree(CODE_DIR)
    azip = zipfile.ZipFile(ZIP_FILE)
    azip.extractall(CODE_DIR)
    azip.close()

    # 文件转移
    master_dir = "./" + CODE_DIR +"/CaiWuTools-master/"
    files = os.listdir(master_dir)
    for file_name in files:
        shutil.move(master_dir + file_name, "./" + file_name)

    # 删除文件
    shutil.rmtree(CODE_DIR)
    os.remove(ZIP_FILE)
    print("更新完毕, 按任意键结束。")

    a = input()
run()