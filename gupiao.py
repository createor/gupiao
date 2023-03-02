#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :  gupiao.py
@Time    :  2023/02/28 17:00:58
@Author  :  createor@github.com
@Version :  1.0
@Desc    :  None
'''

from PyQt5.QtWidgets import QTableView, QAbstractItemView, QMessageBox, QHBoxLayout, QVBoxLayout, QApplication, QMainWindow, QDesktopWidget, QLabel, QAction, QToolBar, qApp, QMenu, QSystemTrayIcon, QDialog, QLineEdit, QPushButton, QCheckBox, QSlider
from PyQt5.QtGui import QIcon, QMouseEvent, QStandardItemModel, QStandardItem, QBrush, QColor
from PyQt5.QtCore import QSize, Qt, QPoint, QPropertyAnimation, QRect
import os, sys
import webbrowser
import configparser
import requests
import win32api, win32con
from apscheduler.schedulers.background import BackgroundScheduler

Headers = {  # 设置固定头部
    "Accept":
    "*/*",
    "Accept-Encoding":
    "gzip, deflate, br",
    "Accept-Language":
    "zh-CN,zh;q=0.9",
    "Host":
    "quote.eastmoney.com",  #
    "Referer":
    "http://quote.eastmoney.com/center/gridlist.html",  #
    "Sec-Ch-Ua":
    "\"Chromium\";v=\"110\", \"Not A(Brand\";v=\"24\", \"Google Chrome\";v=\"110\"",
    "Sec-Ch-Ua-Mobile":
    "?0",
    "Sec-Ch-Ua-Platform":
    "\"Windows\"",
    "Sec-Fetch-Dest":
    "empty",
    "Sec-Fetch-Mode":
    "cors",
    "Sec-Fetch-Site":
    "same-origin",
    "User-Agent":
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36"
}


def search(name_or_code: str, count: int = 5) -> any:
    '''
    根据名称/代码搜索股票
    @param name_or_code:名称或者代码
    @param count:每次查询数量
    @return bool,list:返回成功/失败,结果/原因
    '''
    result = []
    Headers["Host"] = "searchadapter.eastmoney.com"
    url = "https://searchadapter.eastmoney.com/api/suggest/get?input= " + name_or_code + "&type=14&count=" + str(
        count)  # 每次查询5个
    try:
        resp = requests.get(url=url, headers=Headers)
        if resp.status_code == 200:
            data = resp.json()
            for item in data["QuotationCodeTable"]["Data"]:
                result.append({
                    "Code": str(item["Code"]),
                    "Name": str(item["Name"]),
                    "Type": str(item["MktNum"])
                })
            return True, result
        else:
            return False, resp.status_code
    except Exception as e:
        return False, e


def compare(src: str) -> str:
    '''
    格式化结果
    '''
    data = str(src).strip()
    if len(data) == 1:
        return "0.0" + data
    if len(data) == 2:
        return "0." + data
    if len(data) > 2:
        if "." in data:
            return data
        else:
            return data[:-2] + "." + data[-2:]


def getData(code: str) -> any:
    '''
    获取股票当前涨跌
    @param code:股票代码
    @return dict:结果
    '''
    result = {
        "status": False,  # 查询是否成功
        "yestday_price": "",  # 昨收
        "new_price": "",  # 当前价格
        "old_price": "",  # 开盘价
        "rate": "",  # 涨幅
        "time": "",  # 更新时间
        "msg": None  # 错误信息
    }
    Headers["Host"] = "push2.eastmoney.com"
    url = "https://push2.eastmoney.com/api/qt/stock/get?secid=" + code + "&fields=f57,f58,f107,f43,f169,f170,f171,f47,f48,f60,f46,f44,f45,f168,f50,f162,f177"  # 获取开盘价
    # 测试:https://push2.eastmoney.com/api/qt/stock/get?secid=1.603660&fields=f57,f58,f107,f43,f169,f170,f171,f47,f48,f60,f46,f44,f45,f168,f50,f162,f177
    try:
        resp = requests.get(url=url, headers=Headers, timeout=5)
        if resp.status_code == 200:
            resp.encoding = "UTF-8"
            data = resp.json()
            result["status"] = True
            result["new_price"] = compare(str(data["data"]["f43"]))
            result["old_price"] = compare(str(data["data"]["f46"]))
            result["yestday_price"] = compare(str(data["data"]["f60"]))
            temp_data = str(data["data"]["f170"])
            if temp_data.startswith("-"):
                result["rate"] = "-" + compare(temp_data[1:])
            else:
                result["rate"] = "+" + compare(temp_data)
        else:
            result["msg"] = resp.status_code
    except ConnectionError:
        result["msg"] = "请求超时"  # 设置错误信息
    except Exception as e:
        result["msg"] = e
    return result


class app(QMainWindow):
    def __init__(self) -> None:
        super(app, self).__init__()
        # 获取当前程序工作路径
        if getattr(sys, "frozen", False):
            self.workdir = os.path.dirname(os.path.abspath(sys.executable))
        elif __file__:
            self.workdir = os.path.dirname(os.path.abspath(__file__))
        # 设置app标题
        title = "股市详情"
        self.setWindowTitle(title)
        # 设置app尺寸
        self._width = 250
        self._height = 150
        self.resize(self._width, self._height)
        # 设置app图标
        self.icon_path = os.path.join(os.path.join(self.workdir, "icon"),
                                      "gupiao.ico")
        self.setWindowIcon(QIcon(self.icon_path))
        # 设置app居中显示
        self._screen = QDesktopWidget().screenGeometry()
        x = (self._screen.width() - self._width) / 2
        y = (self._screen.height() - self._height) / 2
        self.move(x, y)
        self.setWindowFlags(Qt.WindowStaysOnTopHint
                            | Qt.FramelessWindowHint
                            | Qt.Tool)  # 去除标题栏
        # self.setStyleSheet("border-radius:15px;")  # 设置圆角

        self.config = configparser.ConfigParser()
        self.config.read(os.path.join(os.path.join(self.workdir, "conf"),
                                      "app.ini"),
                         encoding="utf-8")

        if str(self.config.get("settings", "istran")) == "1":
            self.setAttribute(Qt.WA_TranslucentBackground)
            self._tran = True
        else:
            self._tran = False

        if str(self.config.get("settings", "ishide")) == "1":
            self._hide = True
        else:
            self._hide = False

        self._startPos = None
        self._endPos = None
        self._tracking = False

        self.moved = False

        self._timer = 3  # 定时任务周期,单位:秒

        self.stock_list = {}  # 股票列表:{股票代码:股票名称}
        self.stock_code = []
        for code, name in self.config.items("gupiao"):
            self.stock_list[str(code)] = str(name)
            self.stock_code.append(code)

        self.__createTray()
        self.__initUI()
        self.__task()

    def __initUI(self, ) -> None:
        '''
        初始化界面
        '''
        # 标题设置
        label_1 = QLabel("股票名", self)
        label_1.resize(25, 15)
        label_1.move(25, 5)
        label_1.setStyleSheet("font-size:8px;font-weight:bold;")
        label_2 = QLabel("涨跌幅", self)
        label_2.resize(25, 15)
        label_2.move(75, 5)
        label_2.setStyleSheet("font-size:8px;font-weight:bold;")
        label_3 = QLabel("现价", self)
        label_3.resize(20, 15)
        label_3.move(120, 5)
        label_3.setStyleSheet("font-size:8px;font-weight:bold;")
        label_4 = QLabel("昨收", self)
        label_4.resize(20, 15)
        label_4.move(160, 5)
        label_4.setStyleSheet("font-size:8px;font-weight:bold;")
        label_5 = QLabel("今开", self)
        label_5.resize(20, 15)
        label_5.move(200, 5)
        label_5.setStyleSheet("font-size:8px;font-weight:bold;")
        # 底部工具栏设置
        addAct = QAction(
            QIcon(os.path.join(os.path.join(self.workdir, "icon"), "add.ico")),
            '添加', self)
        addAct.triggered.connect(self.__add)  # 绑定事件
        setAct = QAction(
            QIcon(
                os.path.join(os.path.join(self.workdir, "icon"),
                             "setting.ico")), '设置', self)
        setAct.triggered.connect(self.__settings)
        helpAct = QAction(
            QIcon(os.path.join(os.path.join(self.workdir, "icon"),
                               "help.ico")), '帮助', self)
        helpAct.triggered.connect(self.__help)
        tb = QToolBar("工具栏")
        tb.addActions([addAct, setAct, helpAct])
        tb.setIconSize(QSize(16, 16))
        tb.setMovable(False)  # 禁止移动
        tb.setStyleSheet("margin-left:20px;")
        self.addToolBar(Qt.ToolBarArea.BottomToolBarArea, tb)
        # 显示界面
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(["股票名", "涨跌幅", "现价", "昨收",
                                              "今开"])  # 设置标题
        self.layout = QTableView(self)
        self.layout.setModel(self.model)
        self.layout.setGeometry(QRect(10, 20, 230, 110))
        self.layout.setStyleSheet("font-size:8px;")
        self.layout.setShowGrid(False)  # 不显示网线
        self.layout.setEditTriggers(
            QAbstractItemView.NoEditTriggers)  # 设置表格不可修改
        self.layout.verticalHeader().hide()  # 隐藏第一列序号
        self.layout.setColumnWidth(0, 55)
        self.layout.setColumnWidth(1, 40)
        self.layout.setColumnWidth(2, 40)
        self.layout.setColumnWidth(3, 40)
        self.layout.setColumnWidth(4, 40)

    def __task(self, ) -> None:
        '''
        定时任务
        '''
        self.s = BackgroundScheduler()
        self.s.add_job(self.__load, 'interval', seconds=self._timer)
        self.s.start()

    def __load(self, ) -> None:
        '''
        加载数据
        '''
        if self.stock_list:
            for stock in self.stock_list.keys():
                current = self.stock_code.index(stock)
                self.layout.setRowHeight(current, 12)
                reqData = getData(
                    str(self.config.get("type", stock)) + "." + stock)
                if reqData["status"]:
                    self.__draw(
                        current, {
                            "Name": self.stock_list[stock],
                            "Rate": reqData["rate"],
                            "NewPrice": reqData["new_price"],
                            "YestdayPrice": reqData["yestday_price"],
                            "OldPrice": reqData["old_price"]
                        })
                else:
                    print(reqData["msg"])

    def __createTray(self, ) -> None:
        '''
        创建任务栏托盘
        '''
        menu = QMenu(self)
        menu.resize(80, 40)
        menu.setStyleSheet("font-size:8px;")
        action_quit = QAction("退出", self, triggered=self.__quit)  # 添加退出选项
        menu.addAction(action_quit)
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(QIcon(self.icon_path))  # 设置图标
        self.tray.setContextMenu(menu)
        self.tray.show()  # 显示

    def __draw(self, index: int, data: dict) -> None:
        '''
        绘制数据
        '''
        item1 = QStandardItem(data["Name"])
        item2 = QStandardItem(data["Rate"])
        # 判断涨跌幅并设置字体颜色
        try:
            if data["Rate"].startswith("+"):
                item2.setForeground(QBrush(QColor(255, 0, 0)))  # 红色
            if data["Rate"].startswith("-"):
                item2.setForeground(QBrush(QColor(0, 255, 0)))  # 绿色
        except:
            pass
        item3 = QStandardItem(data["NewPrice"])
        item4 = QStandardItem(data["YestdayPrice"])
        item5 = QStandardItem(data["OldPrice"])
        self.model.setItem(index, 0, item1)
        self.model.setItem(index, 1, item2)
        self.model.setItem(index, 2, item3)
        self.model.setItem(index, 3, item4)
        self.model.setItem(index, 4, item5)
        self.layout.viewport().update()  # 更新界面

    def __add(self, ) -> None:
        '''
        '''
        self.addDialog = QDialog()
        self.addDialog.resize(250, 270)
        self.addDialog.setWindowTitle("添加")
        self.addDialog.setWindowFlag(Qt.Tool)
        self.addDialog.setWindowIcon(
            QIcon(os.path.join(os.path.join(self.workdir, "icon"), "add.ico")))
        self.search_word = QLineEdit(self.addDialog)  # 搜索栏
        self.search_word.resize(180, 18)
        self.search_word.move(10, 10)
        self.search_word.setStyleSheet("font-size:8px;")
        self.search_btn = QPushButton("搜索", self.addDialog)  # 搜索按钮
        self.search_btn.resize(30, 18)
        self.search_btn.move(200, 10)
        self.search_btn.clicked.connect(self.__search)
        self.search_btn.setStyleSheet("font-size:8px;")
        # 展示搜索结果
        self.show_search = QVBoxLayout(self.addDialog)
        self.show_search.setGeometry(QRect(0, 40, 230, 220))
        self.show_search.addStretch(1)
        self.show_search.setSpacing(3)
        self.addDialog_x = self.__getPos(250)
        self.addDialog.move(self.addDialog_x, 60)
        self.addDialog.exec()

    def __getPos(self, w) -> int:
        '''
        获取位置
        '''
        pos = self.frameGeometry().topRight()
        if pos.x() + w > self._screen.width():  # 出现再左边
            return pos.x() - self._width - w - 5
        else:
            return pos.x() + 5

    def __search(self, ) -> None:
        '''
        搜索股票
        '''
        keyword = str(self.search_word.text()).strip()
        item_list = list(range(self.show_search.count()))
        item_list.reverse()
        for i in item_list:
            item = self.show_search.itemAt(i)
            self.show_search.removeItem(item)
            if item.widget():
                item.widget().deleteLater()
            if item.layout():
                item_list_ = list(range(item.count()))
                item_list_.reverse()
                for j in item_list_:
                    item_ = item.itemAt(j)
                    item.removeItem(item_)
                    if item_.widget():
                        item_.widget().deleteLater()
        if len(keyword) != 0:
            ok, data = search(keyword)
            if ok:
                for d in data:
                    child_search = QHBoxLayout()
                    child_search.setContentsMargins(20, 0, 20, 0)
                    label = QLabel(d["Name"] + "(" + d["Code"] + ")")
                    label.setStyleSheet("font-size:8px;")
                    btn = QPushButton("添加")
                    btn.setStyleSheet("font-size:8px;")
                    btn.setChecked(True)
                    btn.clicked.connect(
                        lambda u_=btn.isChecked, c_=d["Code"], n_=d["Name"], t_
                        =d["Type"]: self.__attach(u_, c_, n_, t_))
                    child_search.addStretch(1)
                    child_search.addWidget(label, 3)
                    child_search.addWidget(btn, 1)
                    self.show_search.addLayout(child_search)

    def __attach(self, checked: bool, code: str, name: str,
                 code_type: str) -> None:
        '''
        更新配置文件
        @param code:股票代码
        @param name:股票名称
        @param code_type:类型
        '''
        if not self.config.has_option("gupiao", code):
            self.config.set("gupiao", code, name)
            self.config.set("type", code, code_type)
            self.__saveConfig()
            self.stock_list[code] = name  # 添加到列表中
            self.stock_code.append(code)
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("提示")
        msg.setWindowFlag(Qt.Tool)
        msg.setText("添加成功")
        msg.setStyleSheet("font-size:10px;")
        msg_pos = self.addDialog.frameGeometry().topLeft()
        msg.move(msg_pos.x() + 50, msg_pos.y() + 80)
        msg.exec_()

    def __help(self, ) -> None:
        '''
        帮助信息
        '''
        url = os.path.join(self.workdir, "readme.html")
        webbrowser.open_new(url)

    def __settings(self, ) -> None:
        '''
        更新配置
        '''
        self.setDialog = QDialog()
        self.setDialog.resize(220, 150)
        self.setDialog.setWindowTitle("设置")
        self.setDialog.setWindowFlag(Qt.Tool)
        self.setDialog.setWindowIcon(
            QIcon(
                os.path.join(os.path.join(self.workdir, "icon"),
                             "setting.ico")))
        # 设置窗口透明度,滑动
        setLabel = QLabel("设置窗体透明度:", self.setDialog)
        setLabel.resize(65, 20)
        setLabel.move(10, 10)
        setLabel.setStyleSheet("font-size:8px;")
        self.opacitySd = QSlider(Qt.Horizontal, self.setDialog)
        self.opacitySd.setMinimum(1)
        self.opacitySd.setMaximum(10)
        self.opacitySd.setSingleStep(1)
        curr_opacity = int(self.config.get("settings", "opacity"))
        if curr_opacity > 10:
            curr_opacity = 10
        self.opacitySd.setValue(curr_opacity)
        self.opacitySd.resize(130, 20)
        self.opacitySd.move(75, 10)
        self.opacitySd.valueChanged.connect(self.__applyOpacity)

        self.isHide = QCheckBox(self.setDialog)  # 是否开启右侧隐藏
        if str(self.config.get("settings", "ishide")) == "1":
            self.isHide.setChecked(True)
        self.isHide.setText("是否开启靠右隐藏")
        self.isHide.resize(100, 20)
        self.isHide.move(10, 40)
        self.isHide.setStyleSheet("font-size:8px;")
        self.isHide.toggled.connect(self.__applyHide)

        self.isTran = QCheckBox(self.setDialog)  # 是否开启背景透明
        if str(self.config.get("settings", "istran")) == "1":
            self.isTran.setChecked(True)
        self.isTran.setText("是否开启背景透明(重启生效)")
        self.isTran.resize(120, 20)
        self.isTran.move(10, 70)
        self.isTran.setStyleSheet("font-size:8px;")
        self.isTran.toggled.connect(self.__applyTran)

        self.isAuto = QCheckBox(self.setDialog)  # 是否开启开机自启动
        if str(self.config.get("settings", "isauto")) == "1":
            self.isAuto.setChecked(True)
        self.isAuto.setText("是否开启开机自启动")
        self.isAuto.resize(100, 20)
        self.isAuto.move(10, 100)
        self.isAuto.setStyleSheet("font-size:8px;")
        self.isAuto.toggled.connect(self.__applyAuto)

        self.setDialog.move(self.__getPos(220), 80)
        self.setDialog.exec()

    def __applyOpacity(self, ) -> None:
        '''
        窗口透明度
        '''
        self.config.set("settings", "opacity", str(self.opacitySd.value()))
        self.__saveConfig()
        self.setWindowOpacity(self.opacitySd.value() / 10)

    def __applyHide(self, ) -> None:
        '''
        右侧隐藏
        '''
        # 修改配置
        if self.isHide.isChecked():
            self.config.set("settings", "ishide", "1")
            self._hide = True
        else:
            self.config.set("settings", "ishide", "0")
            self._hide = False
        self.__saveConfig()

    def __applyTran(self, ) -> None:
        '''
        背景透明
        '''
        # 修改配置
        if self.isTran.isChecked():
            self.config.set("settings", "istran", "1")
            self._tran = True
        else:
            self.config.set("settings", "istran", "0")
            self._tran = False
        self.__saveConfig()

    def __applyAuto(self, ) -> None:
        '''
        开机自启动
        '''
        # 修改配置
        regName = 'gupiao'
        keyName = r'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
        if self.isAuto.isChecked():
            exe_path = os.path.join(self.workdir, "gupiao.exe")  # 程序路径
            try:
                # 注册表添加
                key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, keyName,
                                          0, win32con.KEY_ALL_ACCESS)
                win32api.RegSetValueEx(key, regName, 0, win32con.REG_SZ,
                                       "\"" + exe_path + "\"")
                win32api.RegCloseKey(key)
                self.config.set("settings", "isauto", "1")
            except:
                pass
        else:
            try:
                # 注册表删除
                key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, keyName,
                                          0, win32con.KEY_ALL_ACCESS)
                win32api.RegDeleteValue(key, regName)
                win32api.RegCloseKey(key)
                self.config.set("settings", "isauto", "0")
            except:
                pass
        self.__saveConfig()

    def __saveConfig(self, ) -> None:
        '''
        保存配置
        '''
        with open(os.path.join(os.path.join(self.workdir, "conf"), "app.ini"),
                  "w+",
                  encoding="utf-8") as f:
            self.config.write(f)

    def enterEvent(self, e: QMouseEvent) -> None:
        '''
        重写鼠标进入事件
        '''
        if self._hide:
            self._hide_or_show("show", e)

    def leaveEvent(self, e: QMouseEvent) -> None:
        '''
        重写鼠标离开事件
        '''
        if self._hide and not self._tran:
            self._hide_or_show("hide", e)

    def _startAnimation(self, w: int, h: int) -> None:
        '''
        动画
        @param w:移动后的x坐标
        @param h:移动后的y坐标
        '''
        animation = QPropertyAnimation(self, b"geometry", self)
        animation.setDuration(200)  # 时间
        new_ops = QRect(w, h, self._width, self._height)
        animation.setEndValue(new_ops)
        animation.start()  # 开始

    def _hide_or_show(self, mode: str, e: QMouseEvent) -> None:
        '''
        判断桌面隐藏还是显示
        @param mode:模式(hide or show)
        '''
        pos = self.frameGeometry().topLeft()  # 左上坐标
        if mode == "show" and self.moved:
            if pos.x() + self._width >= self._screen.width():  # 右侧显示
                self._startAnimation(self._screen.width() - self._width - 12,
                                     pos.y())
                self.moved = False
        if mode == "hide":
            if pos.x() + self._width >= self._screen.width():  # 右侧隐藏
                self._startAnimation(self._screen.width() - 25, pos.y())
                self.moved = True

    def mouseMoveEvent(self, e: QMouseEvent) -> None:
        '''
        重写移动事件
        '''
        if self._tracking:
            self._endPos = e.pos() - self._startPos
            self.move(self.pos() + self._endPos)
            if self._tran:
                self._hide_or_show("hide", e)

    def mousePressEvent(self, e: QMouseEvent) -> None:
        '''
        重写鼠标左键按下事件
        '''
        if e.button() == Qt.LeftButton:
            self._startPos = QPoint(e.x(), e.y())
            self._tracking = True

    def mouseReleaseEvent(self, e: QMouseEvent) -> None:
        '''
        重写鼠标左键释放事件
        '''
        if e.button() == Qt.LeftButton:
            self._tracking = False
            self._startPos = None
            self._endPos = None

    def __quit(self, ) -> None:
        '''
        退出app
        '''
        self.tray.setVisible(False)
        qApp.quit()


if __name__ == '__main__':
    # 运行app
    app_ = QApplication(sys.argv)
    windows = app()
    windows.show()
    sys.exit(app_.exec_())
