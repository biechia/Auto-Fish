from PyQt5.QtWidgets import QApplication
import win32gui, win32com, win32com.client, sys
import pyautogui as pag
from ctypes import CDLL, windll
from numpy import array,uint8,ndarray
import cv2
import aircv as ac
from time import strftime, localtime, sleep


class Mouse:
    def __init__(self):
        self.ghub = CDLL(r'./ghub_mouse.dll')
        if not self.ghub.mouse_open():
            print('未安装ghub或者lgs驱动!!!')
    def move(self, x, y):
        cursor = pag.position()
        x -= cursor.x
        y -= cursor.y
        self.ghub.moveR(x, y)
    def click(self):
        self.ghub.press(1)
        sleep(0.1)
        self.ghub.release(1)

class Screen:
    def __init__(self,win_title=None,win_class=None,hwnd=None) -> None:
        self.app = QApplication(['WindowCapture'])
        self.screen = QApplication.primaryScreen()
        self.bind(win_title,win_class,hwnd)
        self.mouse = Mouse()
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
        template = cv2.imread(templateSrc)
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
        self.mouse.move(x, y)
        self.mouse.click()
        print(f'{strftime("%H:%M:%S", localtime())} 点击坐标: ({x}, {y})')



def main():
    print(f'{strftime("%H:%M:%S", localtime())} [5] 秒后开始执行, 请切换到游戏窗口')
    sleep(5)
    print(f'{strftime("%H:%M:%S", localtime())} 开始操作')

    screen = Screen(win_title='Lost Saga in Timegate - Client')

    while True:

        screen.capture()
        if screen.find('assets/close1.png'): # [关闭(ESC)] 按钮存在
            pass
        elif screen.find('assets/close2.png'): # [关闭(SPACE)] 按钮存在
            pass
        elif screen.find('assets/reward.png'): # [收到奖励] 按钮存在
            pass
        elif screen.find('assets/sell_other.png'): # [贩卖其他道具] 按钮存在
            pass
        elif screen.find('assets/sell_single.png'): # [个别贩卖] 按钮存在
            pass
        elif screen.find('assets/close3.png'): # [×] 按钮存在
            pass
        elif screen.find('assets/sell.png'): # [出售钓鱼产品] 按钮存在
            position = screen.position
            if screen.find('assets/empty_item.png'): # 物品栏为空
                if screen.find('assets/start.png'): # [钓鱼] 按钮存在
                    pass
                else: # 正在钓鱼 -> 等待
                    print(f'{strftime("%H:%M:%S", localtime())} 物品栏为空, 等待 [10] 秒...')
                    sleep(10)
                    continue
            else: # 物品栏不为空
                screen.position = position
        elif screen.find('assets/fish.png'): # [钓鱼] 按钮存在
            pass
        else: # []
            print(f'{strftime("%H:%M:%S", localtime())} 未检测到按钮, [5] 秒后重试')
            sleep(5)
            continue
        screen.focus()
        screen.click()

        # 标出位置
        # screen.pix.save('screen.png')
        # target = cv2.imread('screen.png')
        # rect = position['rectangle']
        # cv2.rectangle(target, (rect[0][0], rect[0][1]), (rect[3][0], rect[3][1]), (0, 0, 220), 2)
        # cv2.imshow("objDetect", target)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()

        print(f'{strftime("%H:%M:%S", localtime())} [5] 秒后执行下一步操作...')
        sleep(5)



if __name__ =='__main__':

    # if not windll.shell32.IsUserAnAdmin():
    #     windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    #     exit()
    
    try:
        main()
    except Exception as e:
        print(e)
        input('回车关闭程序......')




    