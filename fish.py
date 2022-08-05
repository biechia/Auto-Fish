from PyQt5.QtWidgets import QApplication
import win32gui, time
import pyautogui as pag
from ctypes import CDLL
from numpy import array,uint8,ndarray
import cv2
import aircv as ac


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
        time.sleep(0.1)
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
    def find(self, templateSrc):
        template = cv2.imread(templateSrc)
        match_result = ac.find_template(self.img, template)
        # print(templateSrc, match_result)
        return match_result if match_result == None or match_result['confidence'] > 0.9 else None
    def click(self, x, y):
        rect = self.getRect()
        x += rect[0]
        y += rect[1] + 26
        self.mouse.move(x, y)
        self.mouse.click()
        time.sleep(0.1)
        self.mouse.move(rect[0], rect[1])
        


if __name__ =='__main__':

    print('3 秒后开始执行, 请切换到游戏窗口')
    time.sleep(3)
    print("开始操作")

    screen = Screen(win_title='Lost Saga in Timegate - Client')

    while True:
        screen.capture()
        position = screen.find('close1.png')
        if position == None:    # [关闭]按钮不存在
            position = screen.find('reward.png')
            if position == None:    # [收到奖励]按钮不存在
                position = screen.find('sell_other.png')
                if position == None:    # [贩卖其他道具]按钮不存在
                    position = screen.find('sell_single.png')
                    if position == None:    # [个别贩卖]按钮不存在
                        position = screen.find('close2.png')
                        if position == None:    # [x]按钮不存在
                            position = screen.find('sell.png')
                            if position == None:    # [出售钓鱼产品]按钮不存在
                                position = screen.find('fish.png')
                            else:
                                item = screen.find('empty_item.png')
                                if item != None:    # 物品为空
                                    position = screen.find('start.png')
                                    if position == None:    # [钓鱼]按钮不存在
                                        print('无可贩卖物品, 等待 10 秒...')
                                        time.sleep(10)
                                        continue
        
        if position == None:
            ghub = Mouse()
            ghub.move(0, 0)
            time.sleep(3)
        x, y = int(position['result'][0]), int(position['result'][1])
        print(f'点击坐标: ({x}, {y})')

        # 标出位置
        # screen.pix.save('screen.png')
        # target = cv2.imread('screen.png')
        # rect = position['rectangle']
        # cv2.rectangle(target, (rect[0][0], rect[0][1]), (rect[3][0], rect[3][1]), (0, 0, 220), 2)
        # cv2.imshow("objDetect", target)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()

        screen.click(x, y)

        print('3 秒后执行下一步操作...')
        time.sleep(3)





    