
# 第一步先import所需模块（rpcxxxapi，xxx为对应项目的名字）
# rpcwpsapi模块为WPS文字项目的开发接口
# rpcwppapi则是WPS演示的
# rpcetapi毫无疑问就是WPS表格的了
# 另外还有common模块，为前三者的公共接口模块，通常不能单独使用

# 调起WPS必需通过createXXXRpcInstance接口，所以导入它是必需的
# 以WPS文字为例
from pywpsrpc.rpcwpsapi import (createWpsRpcInstance, wpsapi)

# use the RpcProxy to make things easy...
from pywpsrpc import RpcProxy
import os
from sys import argv

# 比如添加一个空白文档

class WordWaterMark:
    
    def __init__(self) -> None:
        pass

    @staticmethod
    def addWaterMark(src_file,des_file,image_path,protect,pwd):
        # 这里仅创建RPC实例
        hr, rpc = createWpsRpcInstance()

        # 注意：
        # WPS开发接口的返回值第一个总是HRESULT（无返回值的除外）
        # 通常不为0的都认为是调用失败（0 == common.S_OK）
        # 可以使用common模块里的FAILED或者SUCCEEDED去判断

        # 通过rpc实例调起WPS进程
        hr, app = rpc.getWpsApplication()
        app.Visible = False
        app.DisplayAlerts = False
        # 打开文档
        hr, doc = app.Documents.Open(src_file)
        doc.Background.Fill.UserTextured(image_path)
        # 文档保护
        if protect == '1':
            doc.Protect(3,True,pwd)
        # 显示背景
        doc.ActiveWindow.View.DisplayBackgrounds=True
        path=os.path.dirname(des_file)
        if not os.path.exists(path):
           os.makedirs(path)
        doc.SaveAs2(des_file)
        app.Quit(wpsapi.wdDoNotSaveChanges)  
        print(des_file)      

if __name__ =='__main__':
    #WordWaterMark.addWaterMark('/opt/1.doc','/opt/water-mark/word/2.docx','/opt/1.png','1','123456')
    WordWaterMark.addWaterMark(argv[1],argv[2],argv[3],argv[4],argv[5])
 
