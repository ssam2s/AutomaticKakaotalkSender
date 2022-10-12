import time
import json
import requests
import collections
import threading
import warnings
from datetime import datetime
from pprint import pprint
from urllib import request
import win32.lib.win32con as win32con
import win32.win32gui as win32gui
import win32.win32api as win32api
import random
warnings.filterwarnings(action='ignore')
collections.Callable = collections.abc.Callable

# ---- Global Variables ----

Thread = []
reqCount = 0
token = ""
keyword1 = ["패스", "패쓰", "써스", "카본", "티그", "알곤", "전기용접", "co2", "씨오투", "카본패스", "카본써스", "tig", "m200", "자동용접", "자동용접사"]
keyword2 = ["배관", "배관사", "덕트", "전기", "칸막이", "배관조공", "포설", "양중"]

# ---- Kakao API ----

# Send Text Message to the Room
def sendText(room, text):
    hwndMain = win32gui.FindWindow(None, room)
    hwndEdit = win32gui.FindWindowEx(hwndMain, None, "RichEdit50W", None)
    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    sendReturn(hwndEdit)

# Send Return Value to Handled Window
def sendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# Open Chat Room
def openRoom(room):
    hwndKakao = win32gui.FindWindow(None, "카카오톡")
    hwndKakao_Edit1 = win32gui.FindWindowEx(hwndKakao, None, "EVA_ChildWindow", None)
    hwndKakao_Edit2_1 = win32gui.FindWindowEx(hwndKakao_Edit1, None, "EVA_Window", None)
    hwndKakao_Edit2_2 = win32gui.FindWindowEx(hwndKakao_Edit1, hwndKakao_Edit2_1, "EVA_Window", None)
    hwndKakao_Edit3 = win32gui.FindWindowEx(hwndKakao_Edit2_2, None, "Edit", None)
    win32api.SendMessage(hwndKakao_Edit3, win32con.WM_SETTEXT, 0, room)
    time.sleep(1)
    sendReturn(hwndKakao_Edit3)
    time.sleep(3)

# ---- Naver Band API ----

# Get Joined Bands List
def GetBands():
    global reqCount
    url = f'https://openapi.band.us/v2.1/bands?access_token={token}'
    reqCount = reqCount + 1
    req = request.Request(url)
    res = request.urlopen(req)
    decoded = res.read().decode('utf8')
    json_dict = json.loads(decoded)
    return json_dict

def GetArticles1():
    global reqCount
    band_key = GetBandKeys()
    locale = 'ko_KR'
    url = f'https://openapi.band.us/v2/band/posts'
    result = []
    dates = []
    for i in range(len(band_key)):
        data = {'access_token': token, 'band_key': band_key[i], 'locale': locale}
        res = requests.get(url, data)
        reqCount = reqCount + 1
        json_dict = res.json()
        for j in range(len(json_dict['result_data']['items'])):
            for x in range(len(keyword1)):
                if keyword1[x] in json_dict['result_data']['items'][j]['content']:
                    if "010" in json_dict['result_data']['items'][j]['content']:
                        d = datetime.fromtimestamp(json_dict['result_data']['items'][j]['created_at'] / 1000)
                        _ = '[' + str(d) + ']\n\n' + json_dict['result_data']['items'][j]['content']
                        if time.strftime("%Y-%m-%d") == str(d).split(' ')[0]:
                            result.append(json_dict['result_data']['items'][j]['content'])
                            dates.append(_)
    result = list(set(result))
    PrintLen(len(result), 1)
    for i in range(len(result)):
        result[i] = dates[i] + result[i]
    return result

def GetArticles2():
    global reqCount
    band_key = GetBandKeys()
    locale = 'ko_KR'
    url = f'https://openapi.band.us/v2/band/posts'
    result = []
    dates = []
    for i in range(len(band_key)):
        data = {'access_token': token, 'band_key': band_key[i], 'locale': locale}
        res = requests.get(url, data)
        reqCount = reqCount + 1
        json_dict = res.json()
        for j in range(len(json_dict['result_data']['items'])):
            for x in range(len(keyword2)):
                if keyword2[x] in json_dict['result_data']['items'][j]['content']:
                    if "010" in json_dict['result_data']['items'][j]['content']:
                        d = datetime.fromtimestamp(json_dict['result_data']['items'][j]['created_at'] / 1000)
                        _ = '[' + str(d) + ']\n\n' + json_dict['result_data']['items'][j]['content']
                        if time.strftime("%Y-%m-%d") == str(d).split(' ')[0]:
                            result.append(json_dict['result_data']['items'][j]['content'])
                            dates.append(_)
    result = list(set(result))
    PrintLen(len(result), 2)
    for i in range(len(result)):
        result[i] = dates[i] + result[i]
    return result

def GetBandKeys():
    json_dict = GetBands()['result_data']['bands']
    result = []
    for i in range(len(json_dict)):
        result.append(json_dict[i]['band_key'])
    return result

def Execution(name, type, limit):
    if type == 1:
        articles = GetArticles1()
    if type == 2:
        articles = GetArticles2() 
    openRoom(name)
    for i in range(len(articles)):
        sendText(name, articles[i])
        PrintSend(type, name)
        time.sleep(55 + random.randint(0, 10))
        if i > limit:
            break

def SubExecution():
    articles = GetArticles2()
    _name = "용접사 일자리 공유및 잡담방 (주말알바)"
    name = "용접에 대해 관심이 있는사람.용접사 자유 소통 일자리"
    openRoom(_name)
    openRoom(name)
    for i in range(len(articles)):
        sendText(name, articles[i])
        sendText(_name, articles[i])
        time.sleep(55 + random.randint(0, 10))
        if i > 3:
            break

def Advertise():
    name = "용접에 대해 관심이 있는사람.용접사 자유 소통 일자리"
    openRoom(name)
    sendText(name, "현장 조공등의 일자리를 24시간 제공(공유)받으려면, 아래의 링크로 입장해주시면 24시간 현장(조공)일자리를 제공 받으실 수 있습니다.\n\nhttps://open.kakao.com/o/giQOrere")
    PrintTime()

def PrintTime():
    hour = time.strftime("%H")
    minute = time.strftime("%M")
    print(hour + ":" + minute + " - Program is Running.")

def PrintLen(len, type):
    hour = time.strftime("%H")
    minute = time.strftime("%M")
    if type == 1:
        print(str(hour) + ":" + str(minute) + " - 용접사 일자리 수 ( " + str(len) + " )")
    elif type == 2:
        print(str(hour) + ":" + str(minute) + " - 조공 일자리 수 ( " + str(len) + " )")

def PrintSend(type, name):
    hour = time.strftime("%H")
    minute = time.strftime("%M")
    if type == 1:
        print(str(hour) + ":" + str(minute) + " - 용접사 일자리 1개 전달 완료 ( " + name + " )")
    elif type == 2:
        print(str(hour) + ":" + str(minute) + " - 조공 일자리 1개 전달 완료 ( " + name + " )")

def CalcReqCount():
    global reqCount
    hour = time.strftime("%H")
    minute = time.strftime("%M")
    print(str(hour) + ":" + str(minute) + " - 현재 일일 API 누적 사용량 ( " + str(reqCount) + " / 1000 )")

if __name__ == '__main__':
    hour = time.strftime("%H")
    minute = time.strftime("%M")
    PrintTime()
    while True:
        _hour = time.strftime("%H")
        _minute = time.strftime("%M")
        if _hour != hour or _minute != minute:
            hour = _hour
            minute = _minute
            PrintTime()
        # Type 1 : 용접사 게시글 수집
        # Type 2 : 조공사 게시글 수집
        if hour == '01' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '02' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '03' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '04' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '05' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '06' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '07' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '08' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '09' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '10' and minute == '00':
            th1 = threading.Thread(target=Advertise, args=())
            th2 = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th1.start()
            th2.start()
            th1.join()
            th2.join()
            CalcReqCount()
        elif hour == '11' and minute == '00':
            th1 = threading.Thread(target=Execution, args=("용접에 대해 관심이 있는사람.용접사 자유 소통 일자리", 1, 60))
            th2 = threading.Thread(target=Execution, args=("용접사 일자리 공유및 잡담방 (주말알바)", 1, 60))
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th1.start()
            th2.start()
            th.start()
            th1.join()
            th2.join()
            th.join()
            CalcReqCount()
        elif hour == '12' and minute == '00':
            th1 = threading.Thread(target=Execution, args=("용접에 대해 관심이 있는사람.용접사 자유 소통 일자리", 1, 60))
            th2 = threading.Thread(target=Execution, args=("용접사 일자리 공유및 잡담방 (주말알바)", 1, 60))
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th1.start()
            th2.start()
            th.start()
            th1.join()
            th2.join()
            th.join()
            CalcReqCount()
        elif hour == '13' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '14' and minute == '00':
            th1 = threading.Thread(target=Execution, args=("용접에 대해 관심이 있는사람.용접사 자유 소통 일자리", 1, 60))
            th2 = threading.Thread(target=Execution, args=("용접사 일자리 공유및 잡담방 (주말알바)", 1, 60))
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th1.start()
            th2.start()
            th.start()
            th1.join()
            th2.join()
            th.join()
            CalcReqCount()
        elif hour == '15' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '16' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '17' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th1 = threading.Thread(target=Advertise, args=())
            th2 = threading.Thread(target=Execution, args=("용접에 대해 관심이 있는사람.용접사 자유 소통 일자리", 1, 60))
            th3 = threading.Thread(target=Execution, args=("용접사 일자리 공유및 잡담방 (주말알바)", 1, 60))
            th.start()
            th1.start()
            th2.start()
            th3.start()
            th.join()
            th1.join()
            th2.join()
            th3.join()
            CalcReqCount()
        elif hour == '18' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th1 = threading.Thread(target=Execution, args=("용접에 대해 관심이 있는사람.용접사 자유 소통 일자리", 1, 60))
            th2 = threading.Thread(target=Execution, args=("용접사 일자리 공유및 잡담방 (주말알바)", 1, 60))
            th.start()
            th1.start()
            th2.start()
            th.join()
            th1.join()
            th2.join()
            CalcReqCount()
        elif hour == '19' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th.start()
            th.join()
            CalcReqCount()
        elif hour == '20' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 60))
            th1 = threading.Thread(target=Execution, args=("용접에 대해 관심이 있는사람.용접사 자유 소통 일자리", 1, 60))
            th2 = threading.Thread(target=Execution, args=("용접사 일자리 공유및 잡담방 (주말알바)", 1, 60))
            th.start()
            th1.start()
            th2.start()
            th.join()
            th1.join()
            th2.join()
            CalcReqCount()
        elif hour == '21' and minute == '00':
            th = threading.Thread(target=Execution, args=("숙식노가다(현장 일자리 24시간 공유방)", 2, 240))
            th.start()
            th.join()
            CalcReqCount()