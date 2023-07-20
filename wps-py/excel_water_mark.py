
# 第一步先import所需模块（rpcxxxapi，xxx为对应项目的名字）
# rpcwpsapi模块为WPS文字项目的开发接口
# rpcwppapi则是WPS演示的
# rpcetapi毫无疑问就是WPS表格的了
# 另外还有common模块，为前三者的公共接口模块，通常不能单独使用

from pywpsrpc.rpcetapi import createEtRpcInstance, etapi
import os
from sys import argv

# 比如添加一个空白文档


class ExcelWaterMark:

    def __init__(self) -> None:
        pass

    @staticmethod
    def addWaterMark(src_file, des_file, image_path, protect, pwd):
        # 这里仅创建RPC实例
        hr, rpc = createEtRpcInstance()

        # 注意：
        # WPS开发接口的返回值第一个总是HRESULT（无返回值的除外）
        # 通常不为0的都认为是调用失败（0 == common.S_OK）
        # 可以使用common模块里的FAILED或者SUCCEEDED去判断

        # 通过rpc实例调起WPS进程
        hr, app = rpc.getEtApplication()
        app.Visible = False
        app.DisplayAlerts = False
        # 打开指定表格文件
        he, workbook = app.Workbooks.Open(src_file)
        sheets=workbook.Worksheets
        for i in range(1,sheets.Count+1):
            sheets[i].SetBackgroundPicture(image_path)

        # 文档保护
        if protect == '1':
           workbook.WritePassword=pwd     
           
        path=os.path.dirname(des_file)
        if not os.path.exists(path):
           os.makedirs(path)
        workbook.SaveAs(des_file)
        workbook.Close()
        app.Quit()
        print(des_file)


if __name__ == '__main__':
    #ExcelWaterMark.addWaterMark('/opt/test.xls','/opt/water-mark/excel/2.xlsx', '/opt/7.png', '1', '123456')
    ExcelWaterMark.addWaterMark(argv[1],argv[2],argv[3],argv[4],argv[5])
