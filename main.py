import requests
import os
import json
import xlwt
import time
import tkinter as tk
import threading
import base64

from tkinter import ttk
from player_score import PlayerScore
from time import sleep

# TODO 使用时添加
# 百度ORC 通用的api key
API_KEY = "GxavEvGtziLkwYpCWHRFNDot"
SECRET_KEY = "n1xB8sBGY8hAldsxqbV8R4ruEvf58jdI"


# 按n切割array
def arr_splits(array, n):
    return [array[i:i + n] for i in range(0, len(array), n)]


def find_img_files(directory, file_dict=None):
    if file_dict is None:
        file_dict = {}
    for dir_path, dir_names, file_names in os.walk(directory):
        jpg_files = [os.path.join(dir_path, f) for f in file_names if f.endswith(('.jpg', '.png', 'jpeg'))]
        if jpg_files:
            file_dict[dir_path] = jpg_files

    return file_dict


# 获取token
def get_access_token():
    """
    使用 AK，SK 生成鉴权签名（Access Token）
    :return: access_token，或是None(如果错误)
    """
    url = "https://aip.baidubce.com/oauth/2.0/token"
    params = {"grant_type": "client_credentials", "client_id": API_KEY, "client_secret": SECRET_KEY}
    return str(requests.post(url, params=params).json().get("access_token"))


# 根据图片文件转换位为base64字符串
def get_img_base64(file_path):
    """
    将图片转换成base64字符串utf-8，不带编码头（data:image/jpeg;base64, ）
    :return: 字符串或None
    """
    # 'C:\\Users\\wangchen1\\Desktop\\百度ORC实践\\弃天.jpg'
    with open(file_path, 'rb') as f:
        image_data = f.read()

    # 若要测试成功，加上编码头去浏览器查看 'data:image/jpeg;base64,'
    # Encode the image data as base64 and remove the header
    return base64.b64encode(image_data).decode('utf-8')


# 根据图片地址，请求orc接口，返回response的text
def get_orc_res(file_path_o, access_token):
    # 测试数据，减少orc请求次数
    # return '{"words_result":[{"words":"壹壹"},{"words":"S267"},{"words":"积分"},{"words":"51062820"},{"words":"积分"},' \
    #        '{"words":"神奇的冬瓜"},{"words":"S263"},{"words":"45670653"},' \
    #        '{"words":"望长安"},{"words":"积分"},{"words":"S4"},{"words":"32552848"},' \
    #        '{"words":"永恒C划水哥"},{"words":"积分"},{"words":"S4"},{"words":"30214445"},' \
    #        '{"words":"小萌新"},{"words":"积分"},{"words":"S267"},{"words":"5776662"},' \
    #        '{"words":"帝江"},{"words":"积分"},{"words":"S258"},{"words":"4083733"},' \
    #        '{"words":"我来也GY"},{"words":"S254"},{"words":"积分"},{"words":"3594500"},' \
    #        '{"words":"收菜的碳酸"},{"words":"S276"},{"words":"积分"},{"words":"3362554"},' \
    #        '{"words":"佛系收菜P"},{"words":"积分"},{"words":"S276"},{"words":"2015089"},' \
    #        '{"words":"仙界核邪￥一一一"},{"words":"积分"},{"words":"S256"},{"words":"1951570"}],' \
    #        '"words_result_num":40,"log_id":1673900975998331160}'

    # 1.加载本地图片，并转化为base64位数据
    base64_data = get_img_base64(file_path_o)

    # 2. 请求接口
    # 2.1 获取百度云接口token
    url = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic?access_token=" + access_token
    payload = {'image': base64_data}
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json'
    }

    return requests.request("POST", url, headers=headers, data=payload).text


# 字符串转换为数组
def data_transfer(data, single_guild_res):
    # 字符串去掉 积分
    data = data.replace('{"words":"积分"},', '')

    # Use json.loads to parse the string into a JSON object
    json_obj = json.loads(data)

    # Access the 'words_result' array from the JSON object
    words_result = json_obj['words_result']
    # 再次分割为数组包含数组
    words_res = arr_splits(words_result, 3)

    for i, item in enumerate(words_res):
        # name, score
        try:
            name = item[0]["words"]
        except IndexError:
            name = ''
        try:
            score = item[2]["words"]
        except IndexError:
            score = ''

        player = PlayerScore(0, name, score)
        single_guild_res.append(player)

    return single_guild_res


# 接受map数据，生成excel
def create_xls(target_path, data):
    # 创建工作簿对象
    work_book = xlwt.Workbook()
    sheet = work_book.add_sheet("积分表")

    # 设置居中对齐
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style = xlwt.XFStyle()  # Create Style
    style.alignment = alignment  # Add Alignment to Style

    c_start = 0
    c_end = 2
    for key, value_list in data.items():
        # 写入一个单元的表头
        # write_merge(a,b,c,d,message)函数将从第a行到第b行的第c列到第d列的单元格合并，并填入内容message
        sheet.write_merge(0, 0, c_start, c_end, key, style=style)
        sheet.write(1, c_start, '序号', style=style)
        sheet.write(1, c_start + 1, '姓名', style=style)
        sheet.write(1, c_start + 2, '积分', style=style)

        for i, item in enumerate(value_list):
            # 推动行
            tmp_row = 1 + i + 1
            sheet.write(tmp_row, c_start, item.rank, style=style)
            sheet.write(tmp_row, c_start + 1, item.name, style=style)
            sheet.write(tmp_row, c_start + 2, item.score, style=style)

        # 下一个单元，推动列
        c_start += 3
        c_end += 3

    # 保存表格
    work_book.save(target_path)


# 方法1
# file_dir = 'D:\\1技术笔记\\考试\\永恒\\img2Excel'
def guild_orc_excel(file_dir, guild_score_dict=None):
    if guild_score_dict is None:
        guild_score_dict = {}

    # 1.仅获取一次百度token，若测试工作量超市，建议改为续时
    access_token = get_access_token()

    # 2.遍历总文件夹或各分文件夹
    file_dict = find_img_files(file_dir)

    # 3.将图片文件遍历，orc处理为对象数组(数组按指定列头积分倒序)
    for key, value_list in file_dict.items():
        # 公会名
        guild_name = str(key).rsplit("\\", 1)[1]

        single_guild_res = []
        for value in value_list:
            orc_res = get_orc_res(value, access_token)
            data_transfer(orc_res, single_guild_res)

            # orc接口有qps = 2，停顿防止报错. 也可以优化sleep > 0.5s的时间
            sleep(1)

        # 积分从大到小排序，并给rank赋值(sort和sorted区别)
        single_guild_res.sort(key=lambda obj: int(obj.score), reverse=True)
        for i, ele in enumerate(single_guild_res, start=1):
            ele.rank = i

        # 加入map
        guild_score_dict[guild_name] = single_guild_res
    # 4.对象map<公会名, [积分数组30]>导出excel
    # 导入excel
    tmp_xls_path = os.path.join(file_dir, '积分表.xls')
    create_xls(tmp_xls_path, guild_score_dict)

    #######################################################################

    for k, v in guild_score_dict.items():
        print(f'--------------------{k}--------------------')
        for value in v:
            print(str(value))

    return tmp_xls_path


# ui调用
def ui_entry(path):
    if len(path) == 0:
        print('请输入地址')

    start_1 = time.perf_counter()  # 返回系统运行时间
    print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())) + '   任务开始请等待')

    # 直接传入参数，后续修改为开发简单ui，由前端校验
    # path = 'D:\\1技术笔记\\考试\\永恒\\img2Excel'

    target_xls_t = guild_orc_excel(path)

    end_1 = time.perf_counter()
    print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
          + ': excel生成完毕。' + (' 耗时：{:.4f}s'.format(end_1 - start_1)) + '      excel路径是： \n' + target_xls_t)
    print('\n *** 谢谢使用，欢迎提出bug并红包求解决！*** \n')


# 前端界面方法
def ui():
    folder_path = input_box.get()
    # ui_entry(folder_path)
    # 异步执行
    thread = threading.Thread(target=ui_entry(folder_path))
    thread.start()

#############################################################


# Create a new Tk root window
root = tk.Tk()

# Set the title of the window
root.title("远征积分转excel工具")

# Set the size of the window
root.geometry("500x200")

# Center the window on the screen
window_width = 500
window_height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = int(screen_height / 2 - window_height / 2)
position_right = int(screen_width / 2 - window_width / 2)
root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

# Create a text entry field for the folder path
input_box = ttk.Entry(root, width=120)
input_box.pack(padx=50, pady=20)


# 创建使用说明的文本框
text = tk.Text(root, height=10, width=40)
description = '使用说明：\n' \
              '1. 请输入文件夹路径。要求：1个文件夹(公会名称)只包含jpg、png、jpeg格式的图片 或者 1个大文件夹包含若干个子文件夹(公会名称)，子文件夹包含指定格式图片 \n' \
              '2. 点击转换后出现未响应莫慌，去你填写的文件夹路径等待若干秒出现excel'
text.insert(tk.END, description)
text.config(state=tk.DISABLED)  # 设置文本框为不可编辑
text.pack()

# 创建一个空标签作为间隔
tk.Label(root, height=1).pack()  # 一个字符大约15px，所以高度设置为50/15=3


# Create a button that calls func1 when clicked
button = tk.Button(root, text='转换', height=30//15, width=70//8, command=ui)
button.pack()

# Start the Tk event loop
root.mainloop()
