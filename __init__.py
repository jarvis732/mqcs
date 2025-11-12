# 方案1：使用相对导入（推荐）
# 方案一：使用绝对导入
from . import MainFunction

if __name__ == '__main__':
    # 定义源数据目录
    dir = '源数据表'
    aimfilpath = '已加工表.xls'
    filelist = []  # 文件路径列表
    MainFunction.initsaveexcel(aimfilpath)
    MainFunction.openfiles(filelist, dir)
    print('打印文件列表：', filelist)  # 打印文件列表
    MainFunction.dealexcel(filelist, aimfilpath)