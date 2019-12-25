import re
import xlrd
from xlutils.copy import copy
import requests
import tkinter as tk
import tkinter.messagebox
import time
import datetime


def str_to_int(date: str):
    if len(date) > 10:
        time_array = time.strptime(date, "%Y-%m-%d %H:%M:%S")
    else:
        time_array = time.strptime(date, "%Y-%m-%d")
    return int(time.mktime(time_array))


def get_request(start, end):
    request_url = "https://edge.allegro.pl/seller/orders?limit=60&sort=-orderDate"
    headers = {
        "accept": "application/vnd.allegro.public.v1+json",
        "accept-encoding": "gzip, deflate, br",
        "accept-language": "zh-CN,zh;q=0.9",
        "content-type": "application/vnd.allegro.public.v1+json",
        "cookie": "_cmuid=uuj999eq-ts5m-5449-psm1-fjtq3m96d47w; ws3=LFKuMps06EuxodMa3Q4oHZa7voFArDEqn; gdpr_permission_given=1; QXLSESSID=c5a941ed7c7ba840fba4c3296bd339c676b6eb27030b7b//02; cartUserId=68771831-ffff-ffff-ffff-ffffffffffffffffffff-ffff-ffff-ffff-ffffffffffff; _gcl_au=1.1.1766029669.1575365323; __gfp_64b=ewRncwFb1NNRuAQy2IUiqm4tdwRiBTdTaWgMfeynCUn.g7; _ga=GA1.2.952287803.1575365330; cd_user_id=16eceaa56c58e9-038d0d78ffa393-32365f08-1fa400-16eceaa56c6976; ahoy_visitor=f8b1f3ac-9fa5-4ef1-8de6-746a4528882d; qeppo_login2=%7B%22welcome%22%3A%22Witaj%2C%22%2C%22id%22%3A%2268771831%22%2C%22username%22%3A%22Amycute%22%2C%22fbConnected%22%3Afalse%2C%22qeppo_hash%22%3A%228T3kGgTA3WWpYbVNJBkgAy6T7A9ocE8Za99rfgv1ni0%3D%22%2C%22isCompany%22%3Atrue%7D; qeppo_priv_cookie=MTE4YgBZB1QACQtTNmEwYw%3D%3D; _bS=0; QXLDATA=e7TIBJjsFnd2Jz5NN6lEvlDQgm%2FGcXcCRavsISII6Hc%3D%23%23%23YVVOHGe9lCqNbast3NnowbSUVk9DKxAfWW125OKDhtNpcy6Ge0uFBi2rbup2TZgw0bpEBeYNZoxhYpD%2B4EgOyuMV; userIdentity=%7B%22id%22%3A%2268771831%22%7D; ws1=MTE4YgBZB1QACQtTNmEwYw%3D%3D; enc_ws1=MTE4YgBZB1QACQtTNmEwYw%3D%3D; dc1=MTE4YgBZB1QACQtTNmEwYw%3D%3D; dc2=ZWIxOHUwfA08BXNielVoeCYzRWx2CXxOPBUMBThkOTk%3D; ws_rec=MTE4YgBZB1QACQtTNmEwYw%3D%3D; _gid=GA1.2.1877124755.1577075437; _gat_UA-2827377-1=1",
        "origin": "https://allegro.pl",
        "referer": "https://allegro.pl/moje-allegro/sprzedaz/zamowienia",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-site",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.88 Safari/537.36",

    }
    data = {
        "limit": "120",
        "sort": "-orderDate",
    }

    responses = requests.get(url=request_url, headers=headers, params=data).json()

    worksheet = xlrd.open_workbook("text.xls", formatting_info=True, encoding_override='utf-8')
    # 获取所有数据表的list
    sheet_name = copy(worksheet)
    sheet = sheet_name.get_sheet(0)
    # 获取长度
    table = worksheet.sheets()[0]
    count = table.nrows
    for text_lists in responses["orders"]:
        try:
            # 下单时间
            bought_at = text_lists["summary"]["paymentLastChanged"]
            time_s = ''.join(re.findall("(.*?)T", bought_at))
            if str_to_int(start) <= str_to_int(time_s) <= str_to_int(end):
                # 订单编号
                order_id = text_lists["id"]
                # 客户ID
                buyer_id = text_lists["buyer"]["id"]
                # 客户的姓
                first_name = text_lists["buyer"]["firstName"]
                # 客户的姓
                last_name = text_lists["buyer"]["lastName"]
                # 客户的邮箱地址
                email = text_lists["buyer"]["email"]
                # 客户的电话
                phone = text_lists["buyer"]["phoneNumber"]
                # 客户详细街道Address1
                street = text_lists["delivery"]["address"]["street"]
                # 客户的收货城市
                city = text_lists["delivery"]["address"]["city"]
                # 客户的省 和市是一样的
                province = text_lists["delivery"]["address"]["city"]
                # 客户的邮箱
                post_code = text_lists["delivery"]["address"]["zipCode"]
                # 客户的收货的国家（简写）
                country_code = text_lists["delivery"]["address"]["countryCode"]
                # 对应渠道SkU
                order_source_sku = text_lists["status"]
                # 数量
                number = text_lists["delivery"]["numberOfPackages"]
                # 订单使用的币种
                currency = text_lists["delivery"]["cost"]["currency"]
                # 单价
                price = text_lists["summary"]["totalToPay"]["amount"]
                # 时间
                time_s = time_s.replace("-", "/")
                sheet.write(count, 0, order_id)
                sheet.write(count, 1, buyer_id)
                sheet.write(count, 4, first_name)
                sheet.write(count, 5, last_name)
                sheet.write(count, 6, email)
                sheet.write(count, 7, phone)
                sheet.write(count, 8, street)
                sheet.write(count, 11, city)
                sheet.write(count, 12, province)
                sheet.write(count, 13, post_code)
                sheet.write(count, 14, country_code)
                sheet.write(count, 15, order_source_sku)
                sheet.write(count, 16, number)
                sheet.write(count, 17, currency)
                sheet.write(count, 18, price)
                sheet.write(count, 19, time_s)
                sheet.write(count, 21, "新平台allegro订单")
                count += 1
        except:
            print("未支付订单，跳过")

    sheet_name.save("output.xls")


def home():
    def download():
        # 以下三行就是获取我们注册时所输入的信息
        starts = start.get()
        ends = end.get()
        if not starts or not ends:
            tkinter.messagebox.showerror(message='开始时间或者结束时间有误！！！')
        else:
            try:
                get_request(starts, ends)
                tkinter.messagebox.showinfo(title='操作成功', message='已经导入到output.xls ')
            except:
                tkinter.messagebox.showerror(message='开始时间或者结束时间有误！！！')
        # window_home.destroy()

    # 定义长在窗口上的窗口
    window_home = tk.Tk()

    photo = tk.PhotoImage(file="background.png")
    ws = window_home.winfo_screenwidth()
    hs = window_home.winfo_screenheight()
    # 计算 x, y 位置
    x = (ws / 2) - (240)
    y = (hs / 2) - (200)
    window_home.geometry('580x400+%d+%d' % (x, y))
    window_home.title('图贸科技')
    window_home.resizable(0, 0)
    tk.Label(window_home, image=photo, compound=tk.CENTER).place(x=0, y=0)
    start = tk.StringVar()
    today = datetime.datetime.now().date()
    start.set(today)
    tk.Label(window_home, bg='#FBF9FA', text='开始时间: ').place(x=170, y=50)
    entry_end = tk.Entry(window_home, textvariable=start)
    entry_end.place(x=240, y=50)

    end = tk.StringVar()
    end.set(today)
    tk.Label(window_home, bg="#FBF9FA", text='结束时间: ').place(x=170, y=90)
    entry_usr_pwd = tk.Entry(window_home, textvariable=end)
    entry_usr_pwd.place(x=240, y=90)

    # 下面的 download
    btn_comfirm_sign_up = tk.Button(window_home, text='下  载', command=download)
    btn_comfirm_sign_up.place(x=240, y=125, relwidth=0.25, relheight=0.08)
    window_home.mainloop()


if __name__ == '__main__':
    home()
