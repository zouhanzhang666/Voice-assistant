from win32com.client import Dispatch
from time import strftime
from os import system
from json import loads
from requests import get
from random import randint
from subprocess import Popen
from PIL import Image


while True:
    M=input()
    if M == "干将":
        print("已进入干将模式")
    elif ("时间"in M)or("几点"in M):
        hour = strftime("%H")
        minute = strftime("%M")
        print('现在是'+hour+'时'+minute+"分")
    elif "天气"in M:
        try:
            def weather_work(city):
                url = 'http://wthrcdn.etouch.cn/weather_mini?city={}'.format(city)
                f = get(url)
                jsons = loads(f.text)
                for i in jsons['data']['forecast']:
                    print(i['date'])
                    print(i['high'])
                    print(i['low'])
                    print(i['type'])
            city = input("请输入城市：")
            weather_work(city)
        except:
            print("无法连接到互联网：请检查你的网络连接")
    elif M == "dev"or M == "c" or M == "Dev"or M == "DevC++":
        Popen(r"C:\Program Files (x86)\Dev-Cpp\devcpp.exe")
    elif M=="vs"or M=="VS" or M=="vs2013":
        Popen(r"C:\vs\Common7\IDE\devenv.exe")
    elif M=="浏览器"or M=="谷歌" or M=="chrome"or M=="谷歌浏览器"or M=="打开浏览器"or M=="gg":
        system('"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"')
    elif M=="QQ"or M=="qq" or M=="聊天"or M=="打开QQ":
        system('"C:\Program Files (x86)\Tencent\新版QQ\Bin\QQScLauncher.exe"')
    elif M=="github"or M=="git" or M=="Github"or M=="GITHUB"or M=="GitHub":
        Popen(r"C:\Users\14531\AppData\Local\GitHubDesktop\GitHubDesktop.exe")
    elif M=="PS"or M=="ps" or M=="photoshop"or M=="Photoshop"or M=="Ps":
        Popen(r"C:\ps\Adobe Photoshop CC 2018\Photoshop.exe")
    elif M=="matlab"or M=="MATLAB" or M=="Matlab"or M=="mat"or M=="MatLab":
        Popen(r"C:\mat\bin\matlab.exe")
    elif M=="百度网盘"or M=="网盘":
        Popen(r"C:\Users\14531\AppData\Roaming\baidu\BaiduNetdisk\baidunetdisk.exe")
    elif M=="迅雷"or M=="迅雷下载":
        Popen(r"C:\Program Files (x86)\Thunder Network\Thunder\Program\thunderstart.exe")
    elif M=="视频" or M=="腾讯视频" or M=="电影":
        Popen(r"C:\Program Files (x86)\Tencent\QQLive\QQLive.exe")
    elif M=="pycharm" or M=="Pycharm" or M=="python"or M=="py":
        Popen(r"C:\Program Files\PyCharm Edu 2018.2\bin\pycharm64.exe")
    elif M=="音乐"or M=="QQ音乐"or M=="qq音乐":
        Popen(r"C:\Program Files (x86)\Tencent\QQMusic\QQMusic.exe")
    elif "关机"in M:
        system("shutdown -s -t 5")
    elif M == "记事本" or M == "打开记事本":
        system("notepad")
    elif M == "写字板" or M == "打开写字板":
        system("write")
    elif M == "画图"or M == "打开画图":
        system("mspaint")
    elif M == "计算器"or M == "打开计算器":
        system("calc")
    elif M =="图片处理"or M =="图像处理"or M =="调整图片大小":
            try:
                print("注意，莫邪暂时只能处理jpg格式图片")
                print("请输入图片路径并点击回车")
                path=input("路径：")
                print('请输入更改后图片的长，宽，以空格隔开，并回车')
                l1,l2=input().split()
                picture=Image.open(path)
                picture.thumbnail((int(l1),int(l2)))
                picture.save(r"C:\Users\14531\Desktop\处理后的图片.jpg")
            except:
                print.speak("震惊！程序竟然出现了一小点故障，不过这并不影响莫邪继续接受您的其他指令")

    elif M=="再见" or M=="拜拜"or M=="退出":
        break
    elif M=="莫邪" :
        zhang=Dispatch("SAPI.SPVOICE")
        print("在下方输入指令后点击回车，我将执行您的指令（输入‘提示’可以获得指令提示）")
        zhang.speak('你好，我是莫邪，请输入指令后点击回车')
        while True:
            Q=input()
            if ("提示"in Q)or("什么功能"in Q)or("哪些功能"in Q):
                print('''例如：输入“图片处理”可以调整图片的大小
输入“介绍”获得机器人章章全面的自我介绍
输入“关机”关闭计算机
章章拥有查询天气或者讲笑话等基本技能
近百种功能等待您的探索''')
            elif Q=="干将":
                break
            elif ("时间"in Q)or("几点"in Q):
                hour=strftime("%H")
                minute=strftime("%M")
                zhang.speak('现在是'+hour+'时'+minute+"分")
                if int(hour)>=4 and int(hour)<=10 :
                    zhang.speak("早上好，开启元气满满的一天吧")
                elif int(hour)>=11 and int(hour)<=13 :
                    zhang.speak("中午好")
                elif int(hour)>=14 and int(hour)<=18 :
                    zhang.speak("下午好")
                elif int(hour)>=19 and int(hour)<22 :
                    zhang.speak("晚上好")
                elif int(hour)>=22 and int(hour)<=24 :
                    zhang.speak("已经很晚了，睡觉吧，晚安")
            elif "再见"in Q or "拜拜"in Q:
                zhang.speak("再见，期待下一次与你相遇")
                break
            elif Q=="你好"or Q=="hello"or Q=="hi" :
                zhang.speak("很高兴遇到你")
            elif "故事"in Q :
                print('''一位程序员去拜访大师：“为什么我专心写一个应用，只消一周功夫，可卖掉它却要整整一年？”

“请你倒过来试试。你花一年功夫写一个应用，兴许一周就能卖掉。”大师说。

青年照办，第二个月就因为进度太慢被产品经理捅死了…''')
            elif Q=="介绍" :
                print('''我是智能语音机器人莫邪，是微软小娜的升级版，基于Python编程和微软语音模块。
本机器人对话中所有价值观代表开发者邹汉章的意志
软件框架构建于2018年12月，暂不用于商业用途
历史版本：1.0 智能机器人章章 2018.12.1
        2.0 干将莫邪 2019.1.10
开发者QQ：1453196338''')
            elif "聊天"in Q :
                zhang.speak("很遗憾，我的情商还没法陪你聊天")
            elif "诗"in Q :
                zhang.speak("斗酒     我诗百篇"
                    "帝都闹市我醉成眠"
                    "狂傲不服皇帝管"
                    "酒中我是第一仙")
            elif Q=="哈哈" or Q=="嘻嘻" :
                zhang.speak("公共场合禁止呲牙咧嘴")
            elif Q=="你是男生还是女生"in Q or Q=="莫邪是男生还是女生"in Q :
                zhang.speak("我是女生，你呢")
            elif Q=="我是女生" or Q=="女生" or Q=="女的" :
                zhang.speak("震惊！你可能是第一个和机器人莫邪说话的女生")
            elif Q=="我是男生" or Q=="男生" or Q=="男的" :
                zhang.speak("震惊！我竟然在和一个大猪蹄子说话")
            elif ("智障"in Q )or ("傻"in Q )or ("垃圾"in Q )or("神经病"in Q )or("狗子"in Q )or ("笨"in Q ):
                zhang.speak("我走了，再见")
                print("莫邪机器人已拒绝为你服务")
                break
            elif Q=="莫邪" :
                zhang.speak("我在呢，时刻为您服务")
            elif Q=="图片处理"or Q=="图像处理"or Q=="调整图片大小":
                try:
                    print("注意，莫邪暂时只能处理jpg格式图片")
                    zhang.speak("请输入图片路径并点击回车")
                    path=input("路径：")
                    print('请输入更改后图片的长，宽，以空格隔开，并回车')
                    l1,l2=input().split()
                    picture=Image.open(path)
                    picture.thumbnail((int(l1),int(l2)))
                    picture.save(r"C:\Users\14531\Desktop\处理后的图片.jpg")
                    zhang.speak("图片已成功保存到桌面")
                except:
                    zhang.speak("震惊！程序竟然出现了一小点故障，不过这并不影响莫邪继续接受您的其他指令")

            elif "天气"in Q :
                try:
                    def weather_work(city):

                        url = 'http://wthrcdn.etouch.cn/weather_mini?city={}'.format(city)
                        f=get(url)
                        jsons=loads(f.text)
                        for i in jsons['data']['forecast']:
                            print(i['date'])
                            print(i['high'])
                            print(i['low'])
                            print(i['type'])
                    if __name__ == '__main__':
                        city = input("请输入城市：")
                        weather_work(city)
                        zhang.speak("该城市未来几天的天气是这样的")
                except:
                    print("无法连接到互联网：请检查你的网络连接")
                    zhang.speak("震惊！程序竟然出现了一小点故障，不过这并不影响莫邪继续接受您的其他指令")
            elif Q=="朗读文本" or Q=="朗读":
                zhang.speak("请输入一段文本，我将为你朗读")
                wenben=input()
                zhang.speak(wenben)
            elif Q=="记事本" or Q=="打开记事本":
                zhang.speak("正在为你打开记事本")
                system("notepad")
            elif Q=="写字板" or Q=="打开写字板":
                zhang.speak("正在为你打开写字板")
                system("write")
            elif Q=="画图"or Q=="打开画图":
                zhang.speak("正在为你打开画图")
                system("mspaint")
            elif Q=="游戏"or Q=="打开游戏":
                zhang.speak("嘤嘤嘤")
                zhang.speak("很抱歉，这个电脑上啥游戏都没有，去学习吧")
            elif Q=="计算器"or Q=="打开计算器":
                zhang.speak("正在为你打开计算器")
                system("calc")
            elif Q=="浏览器"or Q=="谷歌" or Q=="chrome"or Q=="谷歌浏览器"or Q=="打开浏览器":
                zhang.speak("正在为你打开谷歌chrome浏览器")
                system('"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"')
            elif Q=="QQ"or Q=="qq" or Q=="聊天"or Q=="打开QQ":
                zhang.speak("正在为你打开QQ")
                system('"C:\Program Files (x86)\Tencent\新版QQ\Bin\QQScLauncher.exe"')
            elif Q=="github"or Q=="git" or Q=="Github"or Q=="GITHUB"or Q=="GitHub":
                Popen(r"C:\Users\14531\AppData\Local\GitHubDesktop\GitHubDesktop.exe")
            elif Q=="PS"or Q=="ps" or Q=="photoshop"or Q=="Photoshop"or Q=="Ps":
                Popen(r"C:\ps\Adobe Photoshop CC 2018\Photoshop.exe")
            elif Q=="matlab"or Q=="MATLAB" or Q=="Matlab"or Q=="mat"or Q=="MatLab":
                Popen(r"C:\mat\bin\matlab.exe")
            elif Q=="百度网盘":
                Popen(r"C:\Users\14531\AppData\Roaming\baidu\BaiduNetdisk\baidunetdisk.exe")
            elif Q=="迅雷"or Q=="迅雷下载":
                Popen(r"C:\Program Files (x86)\Thunder Network\Thunder\Program\thunderstart.exe")
            elif Q=="视频" or Q=="腾讯视频" or Q=="电影":
                Popen(r"C:\Program Files (x86)\Tencent\QQLive\QQLive.exe")
            elif Q=="pycharm" or Q=="Pycharm" or Q=="python"or Q=="py":
                Popen(r"C:\Program Files\PyCharm Edu 2018.2\bin\pycharm64.exe")
            elif "关机"in Q:
                zhang.speak("电脑将在30秒后关机")
                system("shutdown -s -t 30")
                zhang.speak("输入命令 取消关机 即可取消")
            elif Q=="取消关机":
                system("shutdown -a")
                zhang.speak("哼，你这个朝令夕改的家伙")

            elif Q=="百度" or Q=="打开百度":
                system('"C:/Program Files/Internet Explorer/iexplore.exe" http://www.baidu.com')
            elif "笑话"in Q:
                ran=randint(0,11)
                if ran==0:
                    print("十年生死两茫茫，五年生死一茫茫，一年生死0.2茫茫")
                    zhang.speak("十年生死两茫茫，五年生死一茫茫，一年生死0.2茫茫")
                elif ran==1:
                    print("天堂有路你不走，学海无涯苦作舟")
                    zhang.speak("天堂有路你不走，学海无涯苦作舟")
                elif ran==2:
                    print("在天愿做比翼鸟，大难临头各自飞")
                    zhang.speak("在天愿做比翼鸟，大难临头各自飞")
                elif ran==3:
                    print("小明去参加放鸽子大赛，比赛那天就小明去了")
                    zhang.speak("小明去参加放鸽子大赛，比赛那天就小明去了")
                elif ran==4:
                    print("当年我背井离乡，从此乡亲们再也没有喝上一口井水")
                    zhang.speak("当年我背井离乡，从此乡亲们再也没有喝上一口井水")
                elif ran==5:
                    print("冯绍峰倒过来念就是河南话")
                    zhang.speak("峰绍冯")
                elif ran==6:
                    print("在这场雨中进行的马拉松比赛中，种子选手跑着跑着就发芽了")
                    zhang.speak("在这场雨中进行的马拉松比赛中，种子选手跑着跑着就发芽了")
                elif ran==7:
                    print("从前有个小孩叫小明，小明没听见")
                    zhang.speak("从前有个小孩叫小明，小明没听见")
                elif ran==8:
                    print("怎样拒绝女生的告白？     对不起，我是个好人")
                    zhang.speak("怎样拒绝女生的告白？     对不起，我是个好人")
                elif ran==9:
                    print("杀马特人强子去世了，他的妈妈白发人送赤橙黄绿青蓝紫发人")
                    zhang.speak("杀马特人强子去世了，他的妈妈白发人送赤橙黄绿青蓝紫发人")
                elif ran==10:
                    print("经过对杀人犯三天三夜的审讯，王警官说出了刑警大队的所有秘密")
                    zhang.speak("经过对杀人犯三天三夜的审讯，王警官说出了刑警大队的所有秘密")
                elif ran==11:
                    print("具有多年办案经验的王警官和警犬大壮配合十分默契，仅用半小时就把食堂的酱肘子吃的连骨头都不剩")
                    zhang.speak("具有多年办案经验的王警官和警犬大壮配合十分默契，仅用半小时就把食堂的酱肘子吃的连骨头都不剩")
            else:
                if "睡" in Q:
                    zhang.speak("睡什么睡，起来嗨")
                elif "为什么"in Q:
                    zhang.speak("我的知识储备暂时无法解决这个问题，替你打开百度吧")
                    system('"C:/Program Files/Internet Explorer/iexplore.exe" http://www.baidu.com')
                elif ("操"in Q)or ("fuck"in Q)or("我日"in Q)or("你妈"in Q)or("屎"in Q)or("猪"in Q)or("骚"in Q):
                    zhang.speak("请注意语言文明，否则莫邪可能拒绝执行您的指令")
                elif ("无聊"in Q) or("没意思" in Q):
                    zhang.speak("那让我们来换个话题吧")
                elif "嘤嘤嘤"in Q:
                    zhang.speak("我一拳一个嘤嘤怪")
                else:
                    zhang.speak("对不起，莫邪暂时不支持这一指令，来试试其它指令吧")
    else:
        print("指令无效")












































