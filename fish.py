from PyQt5.QtWidgets import QApplication
import win32gui, win32com, win32com.client, sys
import pyautogui as pag
from ctypes import CDLL
from numpy import array, uint8, ndarray
import cv2
import aircv as ac
from time import strftime, localtime, sleep
from cnocr import CnOcr


class Ghub:
    def __init__(self):
        self.ghub = CDLL(r'./ghub_device.dll')
        if not self.ghub.device_open():
            print('未安装ghub或者lgs驱动!!!')
    def move(self, x, y):
        cursor = pag.position()
        x -= cursor.x
        y -= cursor.y
        self.ghub.moveR(x, y)
    def click(self, delay=0.1):
        self.ghub.mouse_down(1)
        sleep(delay)
        self.ghub.mouse_up(1)
        sleep(delay)
    def key_down(self, keyCode: str):
        self.ghub.key_down(keyCode.encode('utf-8'))
    def key_up(self, keyCode: str):
        self.ghub.key_up(keyCode.encode('utf-8'))
    def key(self, keyCode: str, delay=0.1):
        self.key_down(keyCode)
        sleep(delay)
        self.key_up(keyCode)
        sleep(delay)


class Screen:
    def __init__(self,win_title=None,win_class=None,hwnd=None) -> None:
        self.app = QApplication(['WindowCapture'])
        self.screen = QApplication.primaryScreen()
        self.bind(win_title,win_class,hwnd)
        self.ghub = Ghub()
    def bind(self, win_title=None,win_class=None,hwnd=None):
        '可以直接传入句柄，否则就根据class和title来查找，并把句柄做为实例属性 self._hwnd'
        if not hwnd: self._hwnd = win32gui.FindWindow(win_class, win_title)
        else: self._hwnd = hwnd
    def capture(self, savename='') -> ndarray:
        '截图方法，在窗口为 1920 x 1080 大小下，最快速度25ms (grabWindow: 17ms, to_cvimg: 8ms)'
        def to_cvimg(pix):
            '将self.screen.grabWindow 返回的 Pixmap 转换为 ndarray，方便opencv使用'
            qimg = pix.toImage()
            temp_shape = (qimg.height(), qimg.bytesPerLine() * 8 // qimg.depth())
            temp_shape += (4,)
            ptr = qimg.bits()
            ptr.setsize(qimg.byteCount())
            result = array(ptr, dtype=uint8).reshape(temp_shape)
            return result[..., :3]
        self.pix = self.screen.grabWindow(self._hwnd)
        self.img = to_cvimg(self.pix)
        if savename: self.pix.save(savename)
        return self.img
    def getRect(self):
        return win32gui.GetWindowRect(self._hwnd)
    def focus(self):
        if self._hwnd != win32gui.GetForegroundWindow():
            print(f'{strftime("%H:%M:%S", localtime())} 切换到游戏窗口')
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(self._hwnd)
            sleep(0.1)
    def find(self, templateSrc):
        template = cv2.imread(f'assets/{templateSrc}')
        match_result = ac.find_template(self.img, template)
        # print(templateSrc, match_result)
        if match_result == None or match_result['confidence'] < 0.9:
            self.position = None
            return False
        else: 
            self.position = int(match_result['result'][0]), int(match_result['result'][1])
            return True
    def click(self, x=-1, y=-1):
        if self.position == None:
            print(f'{strftime("%H:%M:%S", localtime())} 未找到点击坐标')
            return
        if x == -1 or y == -1:
            x = self.position[0]
            y = self.position[1]
        rect = self.getRect()
        x += rect[0]
        y += rect[1] + 26
        self.focus()
        self.ghub.move(x, y)
        self.ghub.click()
        print(f'{strftime("%H:%M:%S", localtime())} 点击坐标: ({x}, {y})')
    def key(self, keyCode):
        self.focus()
        self.ghub.key(keyCode)

def sellAuto():
    print(f'{strftime("%H:%M:%S", localtime())} 请把背包翻到第一页, [5] 秒后开始执行程序...')
    sleep(5)
    print(f'{strftime("%H:%M:%S", localtime())} 开始操作')

    screen = Screen(win_title='Lost Saga in Timegate - Client')

    while True:
        screen.capture()

        if screen.find('close1.png'): # [关闭(ESC)] 按钮存在
            pass
        elif screen.find('close2.png'): # [关闭(SPACE)] 按钮存在
            pass
        elif screen.find('reward.png'): # [收到奖励] 按钮存在
            pass
        elif screen.find('calculate.png'): # 自动贩卖计算窗口存在
            img = screen.img
            pos = screen.position
            cropped = img[pos[1] + 35: pos[1] + 60, pos[0] - 95: pos[0] - 50]  # 裁剪坐标为[y0:y1, x0:x1]
            ocr = CnOcr()
            out = ocr.ocr_for_single_line(cropped)
            text = out['text'].replace(' ', '')
            num1, num2 = text.split('+')
            result = str(int(num1) + int(num2))
            print(f'{strftime("%H:%M:%S", localtime())} 计算结果为 [{result}]')
            for i in result:
                screen.key(i)
            screen.ghub.key('enter')
            print(f'{strftime("%H:%M:%S", localtime())} 已输入计算结果, 等待 [60] 秒...')
            sleep(60)
            continue
        elif screen.find('sell_auto_stop.png'): # [停止自动贩卖] 按钮存在
            print(f'{strftime("%H:%M:%S", localtime())} 正在自动贩卖, 等待 [60] 秒...')
            sleep(60)
            continue
        elif screen.find('sell_auto.png'): # [自动贩卖] 按钮存在
            pass
        elif screen.find('friend_ask.png'): # 好友申请窗口存在
            screen.find('friend_accept.png')
        elif screen.find('close3.png'): # [×] 按钮存在
            pass
        elif screen.find('sell.png'): # [出售钓鱼产品] 按钮存在
            position = screen.position
            if screen.find('empty_item_10.png'): # 物品栏未满
                if screen.find('start.png'): # [钓鱼] 按钮存在
                    pass
                else: # 正在钓鱼 -> 等待
                    print(f'{strftime("%H:%M:%S", localtime())} 物品栏未满, 等待 [60] 秒...')
                    sleep(60)
                    continue
            else: # 物品栏已满
                screen.position = position
        elif screen.find('fish.png'): # 钓鱼图形按钮存在
            pass
        else: # 未找到任何按钮
            print(f'{strftime("%H:%M:%S", localtime())} 未检测到按钮, [5] 秒后重试')
            sleep(5)
            continue
        screen.click()
        
        # # 标出位置
        # screen.pix.save('screen.png')
        # target = cv2.imread('screen.png')
        # rect = screen.position
        # cv2.circle(img=target, center=(rect[0], rect[1]), radius=60, color=(0, 0, 255), thickness=2)
        # cv2.imshow("objDetect", target)
        # cv2.waitKey()
        # cv2.destroyAllWindows()

        print(f'{strftime("%H:%M:%S", localtime())} [5] 秒后执行下一步操作...')
        sleep(5)

def sellSingle():
    print(f'{strftime("%H:%M:%S", localtime())} [5] 秒后开始执行程序...')
    sleep(5)
    print(f'{strftime("%H:%M:%S", localtime())} 开始操作')

    screen = Screen(win_title='Lost Saga in Timegate - Client')

    while True:
        screen.capture()

        if screen.find('close1.png'): # [关闭(ESC)] 按钮存在
            pass
        elif screen.find('close2.png'): # [关闭(SPACE)] 按钮存在
            pass
        elif screen.find('reward.png'): # [收到奖励] 按钮存在
            pass
        elif screen.find('sell_other.png'): # [贩卖其他道具] 按钮存在
            pass
        elif screen.find('sell_single.png'): # [个别贩卖] 按钮存在
            pass
        elif screen.find('friend_ask.png'): # 好友申请窗口存在
            screen.find('friend_accept.png')
        elif screen.find('close3.png'): # [×] 按钮存在
            pass
        elif screen.find('sell.png'): # [出售钓鱼产品] 按钮存在
            position = screen.position
            if screen.find('empty_item.png'): # 物品栏为空
                if screen.find('start.png'): # [钓鱼] 按钮存在
                    pass
                else: # 正在钓鱼 -> 等待
                    print(f'{strftime("%H:%M:%S", localtime())} 物品栏为空, 等待 [10] 秒...')
                    sleep(10)
                    continue
            else: # 物品栏不为空
                screen.position = position
        elif screen.find('fish.png'): # 钓鱼图形按钮存在
            pass
        else: # 未找到任何按钮
            print(f'{strftime("%H:%M:%S", localtime())} 未检测到按钮, [5] 秒后重试')
            sleep(5)
            continue
        screen.click()

        # 标出位置
        # screen.pix.save('screen.png')
        # target = cv2.imread('screen.png')
        # rect = screen.position
        # cv2.circle(img=target, center=(rect[0], rect[1]), radius=60, color=(0, 0, 255), thickness=2)
        # cv2.imshow("objDetect", target)
        # cv2.waitKey()
        # cv2.destroyAllWindows()

        print(f'{strftime("%H:%M:%S", localtime())} [5] 秒后执行下一步操作...')
        sleep(5)

def main():
    type = int(input('选择出售方式:\n[1] 自动贩卖: 背包满了之后执行一次自动贩卖(可以干别的事)\n[2] 个别贩卖: 背包不为空则一直执行个别贩卖(适合完全挂机)\n请输入编号: '))
    if type == 1:
        sellAuto()
    elif type == 2:
        sellSingle()
    else:
        print('请输入合法编号!')
        main()

if __name__ =='__main__':

    # if not windll.shell32.IsUserAnAdmin():
    #     windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    #     exit()

    try:
        main()
    except Exception as e:
        print(e)
        input('回车关闭程序...')




    