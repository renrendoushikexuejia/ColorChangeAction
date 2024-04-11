from signal import signal
import sys, os
import json
from threading import Timer
import pyautogui, time
import win32gui,win32api,win32con,winreg
from PyQt5.QtGui import QPixmap, QColor
from PyQt5.QtWidgets import QMainWindow,QApplication,QWidget,QMessageBox,QFileDialog,QColorDialog,QTableWidgetItem
from PyQt5.QtWidgets import QAbstractItemView
from PyQt5.QtCore import pyqtSignal
from ctypes import windll
#导入你写的界面类
from Ui_ColorChangeAction import Ui_ColorChangeAction  
from Ui_Config import Ui_Config

#定义全局变量
pyautogui.FAILSAFE = True
ISRUN = 1   #1是执行  0是停止
ghwndDict = dict()
gpointDict = dict()



#定义全局函数
#获得所有打开的窗口的句柄和名称，存在ghwndDict。在win32gui.EnumWindows(getHwnd, 0)中调用
def getHwnd(hwnd, mouse):
    if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        ghwndDict.update({hwnd:win32gui.GetWindowText(hwnd)})

#获得桌面路径
def getDesktopPath():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key,"Desktop")[0]

# 获取x,y位置像素颜色
def getColor(x, y):
    gdi32 = windll.gdi32
    user32 = windll.user32
    hdc = user32.GetDC(None)  # 获取颜色值
    pixel = gdi32.GetPixel(hdc, x, y)  # 提取RGB值
    r = pixel & 0x0000ff
    g = (pixel & 0x00ff00) >> 8
    b = pixel >> 16
    return {'r':r, 'g':g, 'b':b}


 #MainWindow class----------------------------------------------------------------------
class MyMainWindow(QMainWindow,Ui_ColorChangeAction): 
    #定义一个信号，用来在start--run--scan中发射消息，更新textBrowser
    signalAppendText = pyqtSignal(str)

    #定义函数
    #刷新cbWindowList
    def mfReFresh(self):    
        #获得所有窗口的句柄和名称，添加到cbWindowList
        self.cbWindowList.clear()
        self.cbWindowList.addItem("选择要绑定的窗口")
        win32gui.EnumWindows(getHwnd, 0)
        for k, t in ghwndDict.items():
            if t != '' and t != 'Microsoft Store' and t != 'Microsoft Text Input Application' \
            and t != 'Windows Shell Experience 主机' and t != 'Program Manager' \
            and t != '设置':               #去除title为空 和无用的句柄
                self.cbWindowList.addItem(str(k) + '->' + t)

    #cbWindowList变动时 刷新显示的句柄和窗口名称
    def mfcbWindowListChange(self):
        if  self.cbWindowList.currentText() == "选择要绑定的窗口" or self.cbWindowList.currentText() == "":
            self.labelHwnd.setText("选择要绑定的窗口")
        else:
            strHwnd, strTitle = self.cbWindowList.currentText().split('->')
            self.labelHwnd.setText( "句柄：" + strHwnd + "  窗口名：" + strTitle)

    def mfdsbChange(self):
        if gpointDict:
            gpointDict['frequency'] =  self.dsbFrequency.value()
        else:
            QMessageBox.about( self, "提示", "请先选择方案")

    #打开方案
    def mfOpenProject(self):
        #打开json文件，把json转换为python对象
        mFileName, mFileFilt = QFileDialog.getOpenFileName( self, "打开方案", getDesktopPath(), "json(*.json)")
        try:
            with open( mFileName, 'r', encoding="utf-8") as tempFile:
                global gpointDict 
                gpointDict = json.loads(tempFile.read())

            if gpointDict:
                self.labelPointNum.setText( str(gpointDict['number']))
                self.leProjectName.setText( gpointDict['name'])
                self.dsbFrequency.setValue( gpointDict['frequency'])

        except:
            #打开QFileDialog之后，不选择文件直接关闭，会抛出异常
            pass


    def mfClear(self):
        self.textBrowser.clear()

    #核心代码**********
    def scan(self, hwnd, saveFileDir):
        #截图保存
        global ISRUN
        img = QApplication.primaryScreen().grabWindow( int(hwnd)).toImage()     #hwnd是正整数，传的参数是字符串
        img.save( saveFileDir)


        for i, k in gpointDict.items():
            if isinstance( k, dict):
                if int(k['px']) < img.width()  and  int(k['py']) < img.height():
                    pixelColor = QColor( img.pixel( int(k['px']), int(k['py'])))  #坐标str转为int
                    r, g, b = pixelColor.red(), pixelColor.green(), pixelColor.blue()
                
                    setR, setG, setB = k['pNow'][1:-1].split(',')
                    if not (r == int(setR) and g == int(setG) and b == int(setB)):
                        
                        signalStr = k['pName'] + "-改变" + " 检测:(" + str(r) + ","+ str(g) + ","+ str(b) + ")" + \
                            " 预设:(" + str(setR) + ","+ str(setG) + ","+ str(setB) + ")" 
                        self.signalAppendText.emit( signalStr)

                        time.sleep(1)
                    else:
                        #如果检测点颜色没有改变，则不做提示
                        time.sleep(1)


                else:
                    #QMessageBox.about( myMW, "提示", "坐标错误，请检查方案")
                    errorStr = "坐标错误，请检查方案 或 被检测窗口是否最小化"
                    self.signalAppendText.emit( errorStr)
                    ISRUN = 0
                    break

                # signalStr = "检测点: " + k['pName'] + " 预设: " + k['pNow']  + 
                # " 检测: (" + str(r) + "," + str(g) + "," + str(b) + ")"
                # self.signalAppendText.emit( signalStr)      #直接在新线程中改变主界面会引发警告，要使用发送信号的方式


    #开始执行
    def start(self):
        #获得执行窗口的句柄
        tempHwnd = self.cbWindowList.currentText().split('->')[0]
        if tempHwnd == "选择要绑定的窗口":
            QMessageBox.about( self, "提示", tempHwnd)
            return
        
        if not gpointDict:
            QMessageBox.about( self, "提示", "请选择方案")
            return

        if gpointDict['number'] == 0:
            QMessageBox.about( self, "提示", "请检查方案")
            return

        #创建保存截图的文件夹
        saveFileDir = getDesktopPath() + "\\" + gpointDict['name'] + "\\" + gpointDict['name'] + ".jpg" 
        saveFolderDir = getDesktopPath() + "\\" + gpointDict['name']
        if not os.path.exists( saveFolderDir):
            os.mkdir( saveFolderDir)


        global ISRUN
        ISRUN = 1
        #在start中定义run，在run中定义Timer， Timer调用run，一直循环，直到点击stop 使全局ISRUN == 0
        def run():
            #print( ISRUN)  #通过ISRUN的值来判断是否执行

            #定义一个Timer， 控制start run stop
            tempTimer = Timer( gpointDict['frequency'], run)
            tempTimer.start()
            if ISRUN == 0:
                tempTimer.cancel()
            else:
                self.scan( tempHwnd, saveFileDir)

        run()


    #停止执行
    def stop( self):
        global ISRUN
        ISRUN = 0


    #接收线程发过来的str，更新textBrowser
    def mfappendText( self, tempStr):
        self.textBrowser.append( tempStr)


    #构造函数
    def __init__(self,parent =None):
        super(MyMainWindow,self).__init__(parent)
        self.setupUi(self)

        #获得所有窗口的句柄和名称，添加到cbWindowList
        self.cbWindowList.addItem("选择要绑定的窗口")
        win32gui.EnumWindows(getHwnd, 0)
        for k, t in ghwndDict.items():
            if t != '' and t != 'Microsoft Store' and t != 'Microsoft Text Input Application' \
            and t != 'Windows Shell Experience 主机' and t != 'Program Manager' \
            and t != '设置':               #去除title为空 和无用的句柄
                self.cbWindowList.addItem(str(k) + '->' + t)

       
        #绑定按键和信号槽
        self.btnRefresh.clicked.connect( self.mfReFresh)
        self.cbWindowList.currentIndexChanged.connect( self.mfcbWindowListChange)
        self.btnOpenProject.clicked.connect( self.mfOpenProject)
        self.btnStart.clicked.connect( self.start)
        self.btnStop.clicked.connect( self.stop)
        self.dsbFrequency.valueChanged.connect( self.mfdsbChange)
        self.btnClear.clicked.connect( self.mfClear)
        self.signalAppendText.connect( self.mfappendText)
    

#ConfigWindow class----------------------------------------------------------------------------
class MyConfigWindow(QMainWindow,Ui_Config): 
    
    #定义槽函数
    #删除所选行
    def mfdeleteRow(self):
        self.twpoint.removeRow( self.twpoint.currentRow())


    #刷新cbWindowList
    def mfReFresh(self):    
    #获得所有窗口的句柄和名称，添加到cbWindowList
        self.cbWindowList.clear()
        self.cbWindowList.addItem("选择要绑定的窗口")
        win32gui.EnumWindows(getHwnd, 0)
        for k, t in ghwndDict.items():
            if t != '' and t != 'Microsoft Store' and t != 'Microsoft Text Input Application' \
            and t != 'Windows Shell Experience 主机' and t != 'Program Manager' \
            and t != '设置':               #去除title为空 和无用的句柄
                self.cbWindowList.addItem(str(k) + '->' + t)

    
    #cbWindowList变动时 刷新显示的句柄和窗口名称
    def mfcbWindowListChange(self):
        if  self.cbWindowList.currentText() == "选择要绑定的窗口" or self.cbWindowList.currentText() == "":
            self.labelHwnd.setText("选择要绑定的窗口")
        else:
            strHwnd, strTitle = self.cbWindowList.currentText().split('->')
            self.labelHwnd.setText( "句柄：" + strHwnd + "  窗口名：" + strTitle)

        
    #添加检测点
    def mfaddPoint(self):
        tempHwnd = self.cbWindowList.currentText().split('->')[0]
        
        if tempHwnd != "选择要绑定的窗口":

            win32gui.SetForegroundWindow( tempHwnd)     #窗口被遮挡时，使窗口前端显示
            
            # win32gui.ShowWindow( tempHwnd, win32con.SW_RESTORE)     #窗口最小化时，使窗口前端显示
            if self.cbShow.currentText() == "全屏":         #分为全屏和窗口，只在全屏时做改变
                win32gui.ShowWindow( tempHwnd, win32con.SW_SHOWMAXIMIZED)
            else:
                win32gui.ShowWindow( tempHwnd, win32con.SW_SHOWNORMAL)


            time.sleep(2)       #延迟执行，这个时间是写在程序里固定的。 点击添加按键后等待数秒，执行获取坐标和颜色操作
            topX, topY =pyautogui.position()
            rgbDict = getColor( topX, topY)
            tempStr = "坐标:("+str(topX)+","+str(topY)+ ")  颜色RGB("+str(rgbDict['r'])+","+str(rgbDict['g'])+","+str(rgbDict['b'])+")"

            try:
                win32gui.SetForegroundWindow( win32gui.FindWindow( None, "方案设置"))   #通过标题名，找到窗口句柄，并前端显示
            except:
                pass
            QMessageBox.about( self, "检测", tempStr)
            

            windowX, windowY = win32gui.ScreenToClient( tempHwnd, (topX, topY)) #把屏幕坐标转为相对窗口左上角的坐标

            mrow = self.twpoint.rowCount()
            self.twpoint.insertRow( mrow)
            self.twpoint.setItem( mrow, 0, QTableWidgetItem("双击修改"))
            self.twpoint.setItem( mrow, 1, QTableWidgetItem(str(windowX)))
            self.twpoint.setItem( mrow, 2, QTableWidgetItem(str(windowY)))
            self.twpoint.setItem( mrow, 3, QTableWidgetItem("("+str(rgbDict['r'])+','+str(rgbDict['g'])+','+str(rgbDict['b'])+")"))
            self.twpoint.setItem( mrow, 4, QTableWidgetItem("双击修改"))
            self.twpoint.setItem( mrow, 5, QTableWidgetItem("双击修改"))
            self.twpoint.setItem( mrow, 6, QTableWidgetItem("双击修改"))

        else:
            QMessageBox.about( self, "提示","请先选择检测点所在窗口")

    #保存方案  btnsave的槽函数
    def mfsave( self):
        if self.lenameConfig.text() != '' and self.twpoint.rowCount() != 0:

            tempDict = dict()
            tempDict['name'] = self.lenameConfig.text()
            tempDict['number'] = self.twpoint.rowCount()
            tempDict['frequency'] = self.dsbfrequencyConfig.value()

            for i in range( 0, tempDict['number']):
                pointDict = dict()

                pointDict['pName'] = self.twpoint.item( i, 0).text()
                pointDict['px'] = self.twpoint.item( i, 1).text()
                pointDict['py'] = self.twpoint.item( i, 2).text()
                pointDict['pNow'] = self.twpoint.item( i, 3).text()
                pointDict['pTo'] = self.twpoint.item( i, 4).text()
                pointDict['pMsg'] = self.twpoint.item( i, 5).text()
                pointDict['pMsgColor'] = self.twpoint.item( i, 6).text()
 
                temp = "point" + str(i)
                tempDict[ temp] = pointDict
            
            #json.dumps(odata,ensure_ascii=False).decode('utf8').encode('gb2312')   不可用
            #直接用json.dumps()， json文件中会出现 斜杠开头的字符 \， 需添加参数ensure_ascii=False
            #indent=4 是格式化文件中的json，向右缩进4个字符
            tempJson = json.dumps( tempDict, indent=4, ensure_ascii=False)

            #把配置方案写入JSON文件
            tempfileDir = getDesktopPath() + "\\" + tempDict['name'] + "\\" +tempDict['name'] + ".json"
            if os.path.exists( tempfileDir):

                with open( tempfileDir, 'w', encoding='utf-8') as f:
                    f.write( tempJson)

            else:
                #要创建目录，然后写入文件
                tempMkdir = getDesktopPath() + "\\" + tempDict['name']
                if not os.path.exists( tempMkdir):
                    os.mkdir( tempMkdir)

                with open( tempfileDir, 'x', encoding='utf-8') as f:
                    f.write( tempJson)

        #点过保存之后，还要刷新myWM界面，关闭当前窗口，更新gpointDict字典
            global gpointDict 
            gpointDict = tempDict
            myMW.leProjectName.setText( gpointDict['name'])
            myMW.labelPointNum.setText( str(gpointDict['number']))
            myMW.dsbFrequency.setValue( gpointDict['frequency'])

            self.close()

        else:
            QMessageBox.about( self, "提示","请添加检测点和填写方案名称")


    def __init__(self,parent =None):
        super(MyConfigWindow,self).__init__(parent)
        self.setupUi(self)

        #获得所有窗口的句柄和名称，添加到cbWindowList
        self.cbWindowList.addItem("选择要绑定的窗口")
        win32gui.EnumWindows(getHwnd, 0)
        for k, t in ghwndDict.items():
            if t != '' and t != 'Microsoft Store' and t != 'Microsoft Text Input Application' \
            and t != 'Windows Shell Experience 主机' and t != 'Program Manager' \
            and t != '设置':               #去除title为空 和无用的句柄
                self.cbWindowList.addItem(str(k) + '->' + t)


        #初始化表格
        tempHeadText = ["点位名","X","Y","当前颜色","变为颜色","消息","消息背景"]
        self.twpoint.setColumnCount( len(tempHeadText))
        self.twpoint.setHorizontalHeaderLabels( tempHeadText)

        #初始化cbShow 显示方式
        self.cbShow.addItem("全屏")
        self.cbShow.addItem("窗口")


        #绑定槽函数
        self.btnDeleteRow.clicked.connect( self.mfdeleteRow)
        self.btnaddPoint.clicked.connect( self.mfaddPoint)
        self.btnRefresh.clicked.connect( self.mfReFresh)
        self.cbWindowList.currentIndexChanged.connect( self.mfcbWindowListChange)
        self.btnsave.clicked.connect( self.mfsave)
        self.btnexit.clicked.connect( self.close)


 #程序入口----------------------------------------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myMW = MyMainWindow()
    myCW = MyConfigWindow()
    
    myMW.show()

    #定义槽函数
    def fopenConfig():
        if gpointDict:
            myCW.lenameConfig.setText( gpointDict['name'])
            myCW.dsbfrequencyConfig.setValue( gpointDict['frequency'])

            #向表格里添加数据
            # myCW.twpoint.clearContents()        这个是清空数据，但是不移除行。再次调用fopenConfig(),就会出现空行。
            for i in range( 0, myCW.twpoint.rowCount()):
                myCW.twpoint.removeRow(0)

            myCW.twpoint.setSelectionBehavior( QAbstractItemView.SelectRows)   #设置表格只能按行选择
            n = 0
            for i, k in gpointDict.items():
                if isinstance( k, dict):

                    myCW.twpoint.insertRow( n)
                    myCW.twpoint.setItem( n, 0, QTableWidgetItem(k['pName']))
                    myCW.twpoint.setItem( n, 1, QTableWidgetItem(str(k['px'])))
                    myCW.twpoint.setItem( n, 2, QTableWidgetItem(str(k['py'])))
                    myCW.twpoint.setItem( n, 3, QTableWidgetItem(k['pNow']))
                    myCW.twpoint.setItem( n, 4, QTableWidgetItem(k['pTo']))
                    myCW.twpoint.setItem( n, 5, QTableWidgetItem(k['pMsg']))
                    myCW.twpoint.setItem( n, 6, QTableWidgetItem(k['pMsgColor']))

                    n = n + 1
            
            myCW.show()
            
        else:
            myCW.show()


    #绑定槽函数
    myMW.btnProjectConfig.clicked.connect( fopenConfig)
    
sys.exit(app.exec_())  