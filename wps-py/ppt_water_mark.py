
# 第一步先import所需模块（rpcxxxapi，xxx为对应项目的名字）
# rpcwpsapi模块为WPS文字项目的开发接口
# rpcwppapi则是WPS演示的
# rpcetapi毫无疑问就是WPS表格的了
# 另外还有common模块，为前三者的公共接口模块，通常不能单独使用

from pywpsrpc.rpcwppapi import createWppRpcInstance,wppapi

from sys import argv

# 比如添加一个空白文档


class PowerpointWaterMark:

    def __init__(self) -> None:
        pass

    @staticmethod
    def addWaterMark(src_file, des_file, image_path, protect, pwd):
        # 这里仅创建RPC实例
        hr, rpc = createWppRpcInstance()

        # 注意：
        # WPS开发接口的返回值第一个总是HRESULT（无返回值的除外）
        # 通常不为0的都认为是调用失败（0 == common.S_OK）
        # 可以使用common模块里的FAILED或者SUCCEEDED去判断

        # 通过rpc实例调起WPS进程
        hr, app = rpc.getWppApplication()
        app.Visible = False
        app.DisplayAlerts = False
        # 打开指定表格文件
        he, ppt = app.Presentations.Open(src_file)
        ppt.SlideMaster.Background.Fill.UserTextured=image_path
        # 文档保护
        if protect == '1':
            ppt.WritePassword = pwd

        ppt.SaveAs(des_file)
        ppt.Close()
        app.Quit()
        print(des_file)


if __name__ == '__main__':
    PowerpointWaterMark.addWaterMark('/opt/新建 PPTX 演示文稿.pptx','/opt/water-mark/ppt/2.pptx', '/opt/7.png', 0, '123456')
    # ExcelWaterMark.addWaterMark(argv[1],argv[2],argv[3],argv[4],argv[5])
