import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QComboBox, QFrame, QMessageBox, QLineEdit, QTextEdit
from PyQt5.QtWidgets import QHBoxLayout, QVBoxLayout, QPushButton, QTableView, QCalendarWidget
from PyQt5.QtCore import QAbstractTableModel, Qt, QDate
from PyQt5.QtGui import QIcon, QFont
import datetime
import pandas as pd


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("NancycyZL")
        self.resize(800, 600)

        vlayout = QVBoxLayout(self)

        label_title = QLabel('众兴文具会员管理系统')
        label_title.setFont(QFont('华文行楷', 36))
        label_title.setAlignment(Qt.AlignCenter)

        hlayout = QHBoxLayout()
        btn1 = QPushButton('添加会员')
        btn1.setFont(QFont('华文行楷', 18))
        btn1.clicked.connect(self.myAddMember)
        btn2 = QPushButton('添加项目')
        btn2.setFont(QFont('华文行楷', 18))
        btn2.clicked.connect(self.myAddItem)
        btn3 = QPushButton('修改信息')
        btn3.setFont(QFont('华文行楷', 18))
        btn3.clicked.connect(self.myChkMod)
        btn4 = QPushButton('查询记录')
        btn4.setFont(QFont('华文行楷', 18))
        btn4.clicked.connect(self.myChkRecord)
        hlayout.addWidget(btn1)
        hlayout.addWidget(btn2)
        hlayout.addWidget(btn3)
        hlayout.addWidget(btn4)

        vlayout.addStretch(2)
        vlayout.addWidget(label_title)
        vlayout.addStretch(1)
        vlayout.addLayout(hlayout)
        vlayout.addStretch(2)

        self.show()

    def myAddItem(self):
        window_list.append(AddItemWindow())

    def myAddMember(self):
        window_list.append(AddMemberWindow())

    def myChkMod(self):
        window_list.append(ChkModWindow())

    def myChkRecord(self):
        window_list.append(ChkRecordWindow())


class AddMemberWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.ID = 0
        self.NAME = ''
        self.PHONE = ''
        self.NOTE = ''

    def initUI(self):
        self.setWindowTitle("添加会员")
        self.resize(400, 300)

        vlayout = QVBoxLayout(self)

        hlayout_1 = QHBoxLayout()
        hlayout_1.addWidget(QLabel('    ID'))
        self.le1 = QLineEdit()
        self.le1.textChanged[str].connect(self.changeID)
        hlayout_1.addWidget(self.le1)

        hlayout_2 = QHBoxLayout()
        hlayout_2.addWidget(QLabel('姓名'))
        self.le2 = QLineEdit()
        self.le2.textChanged[str].connect(self.changeNAME)
        hlayout_2.addWidget(self.le2)

        hlayout_3 = QHBoxLayout()
        hlayout_3.addWidget(QLabel('电话'))
        self.le3 = QLineEdit()
        self.le3.textChanged[str].connect(self.changePHONE)
        hlayout_3.addWidget(self.le3)

        hlayout_4 = QHBoxLayout()
        hlayout_4.addWidget(QLabel('备注'))
        self.le4 = QLineEdit()
        self.le4.textChanged[str].connect(self.changeNOTE)
        hlayout_4.addWidget(self.le4)

        btn = QPushButton('确定')
        btn.clicked.connect(self.confirm)

        vlayout.addStretch(3)
        vlayout.addLayout(hlayout_1)
        vlayout.addStretch(1)
        vlayout.addLayout(hlayout_2)
        vlayout.addStretch(1)
        vlayout.addLayout(hlayout_3)
        vlayout.addStretch(1)
        vlayout.addLayout(hlayout_4)
        vlayout.addStretch(2)
        vlayout.addWidget(btn)
        vlayout.addStretch(3)

        self.show()

    def changeID(self, n):
        self.ID = n

    def changeNAME(self, name):
        self.NAME = name

    def changePHONE(self, phone):
        self.PHONE = phone

    def changeNOTE(self, note):
        self.NOTE = note

    def clear_input(self):
        self.le1.clear()
        self.le2.clear()
        self.le3.clear()
        self.le4.clear()

    def valid_input(self, id, phone):
        try:
            id = int(id)
        except ValueError:
            msg_error = QMessageBox(QMessageBox.Warning, '错误', '请输入有效的ID(整数）', QMessageBox.Close, self)
            msg_error.show()
            return False

        try:
            phone = int(phone)
        except ValueError:
            msg_error = QMessageBox(QMessageBox.Warning, '错误', '请输入有效的电话号码', QMessageBox.Close, self)
            msg_error.show()
            return False

        existing_id = member_df['ID'].values
        if id in existing_id:
            msg_error = QMessageBox(QMessageBox.Warning, '错误', '该ID以存在！', QMessageBox.Close, self)
            msg_error.show()
            return False
        else:
            return True


    def confirm(self):
        if self.valid_input(self.ID, self.PHONE):
            new_info = pd.DataFrame({'ID':int(self.ID), '姓名':self.NAME, '电话':int(self.PHONE), '积分':0, '余额':0, '备注':self.NOTE}, index=[0])
            global member_df
            new_df = member_df.append(new_info).sort_values(by=['ID']).reset_index(drop=True)
            new_df.to_excel('会员信息.xlsx', sheet_name='Sheet1', index=False)
            member_df = new_df
            msg_confirm = QMessageBox(QMessageBox.Information,'确认','已添加{}号会员：{}'.format(self.ID,self.NAME),QMessageBox.Ok,self)
            msg_confirm.show()
            self.clear_input()


class AddItemWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.AMOUNT = 0
        self.initUI()

    def initUI(self):
        self.setWindowTitle("添加项目")
        self.resize(400, 300)

        global member_df
        global item_df

        id_list = member_df['ID'].values
        id_list_str = [str(id) for id in id_list]
        name_list = member_df['姓名'].values

        hlayout1 = QHBoxLayout()
        lbl_id = QLabel('    ID')
        lbl_id.setFixedWidth(50)
        self.combox_id = QComboBox()
        self.combox_id.addItem('请选择ID')
        self.combox_id.addItems(id_list_str)
        self.combox_id.activated[str].connect(self.myActivateID)
        lbl_name = QLabel('    姓名')
        lbl_name.setFixedWidth(50)
        self.combox_name = QComboBox()
        self.combox_name.addItem('请选择名字')
        self.combox_name.addItems(name_list)
        self.combox_name.activated[str].connect(self.myActivateNAME)
        btn_reset = QPushButton('重置选项')
        btn_reset.clicked.connect(self.resetChoices)
        hlayout1.addWidget(lbl_id)
        hlayout1.addWidget(self.combox_id)
        hlayout1.addWidget(lbl_name)
        hlayout1.addWidget(self.combox_name)
        hlayout1.addStretch(3)
        hlayout1.addWidget(btn_reset)

        hlayout2 = QHBoxLayout()
        self.le_amount = QLineEdit()
        self.le_amount.textChanged[str].connect(self.setAMOUNT)
        hlayout2.addWidget(QLabel('消费金额'))
        hlayout2.addWidget(self.le_amount)

        hlayout3 = QHBoxLayout()
        self.te_note = QTextEdit()
        hlayout3.addWidget(QLabel('备       注'))
        hlayout3.addWidget(self.te_note)

        btn_confirm = QPushButton('确定')
        btn_confirm.clicked.connect(self.inputItem)

        vlayout = QVBoxLayout(self)
        vlayout.addStretch(3)
        vlayout.addLayout(hlayout1)
        vlayout.addStretch(1)
        vlayout.addLayout(hlayout2)
        vlayout.addStretch(1)
        vlayout.addLayout(hlayout3)
        vlayout.addStretch(1)
        vlayout.addWidget(btn_confirm)
        vlayout.addStretch(3)

        self.show()

    def myActivateID(self,id):
        # print(id)
        global member_df
        self.combox_name.clear()
        if id != '请选择ID':
            name = member_df.loc[member_df['ID']==int(id)]['姓名'].values[0]
            self.combox_name.addItem(name)
        else:
            self.combox_name.addItem('请选择名字')
            self.combox_name.addItems(member_df['姓名'].values)

    def myActivateNAME(self,name):
        self.combox_id.clear()
        self.combox_id.addItem('请选择ID')
        global member_df
        ids_int = member_df.loc[member_df['姓名'] == name]['ID'].values
        # print(ids_int)
        ids_choices = [str(id) for id in ids_int]
        self.combox_id.addItems(ids_choices)

    def resetChoices(self):
        global member_df
        self.combox_id.clear()
        self.combox_id.addItem('请选择ID')
        self.combox_id.addItems([str(id) for id in member_df['ID'].values])
        self.combox_name.clear()
        self.combox_name.addItem('请选择名字')
        self.combox_name.addItems(member_df['姓名'].values)

    def clear_input(self):
        self.le_amount.clear()
        self.te_note.clear()

    def setAMOUNT(self, amount):
        self.AMOUNT = amount

    def valid_input(self, id, name, amount):
        if id == '请选择ID':
            msg_id = QMessageBox(QMessageBox.Warning, '错误', '请选择有效ID', QMessageBox.Ok, self)
            msg_id.show()
            return False
        if name == '请选择名字':
            msg_name = QMessageBox(QMessageBox.Warning, '错误', '请选择有效名字', QMessageBox.Ok, self)
            msg_name.show()
            return False
        if amount == 0 or amount == '0':
            msg_amount = QMessageBox(QMessageBox.Warning, '错误', '请输入金额', QMessageBox.Ok, self)
            msg_amount.show()
            return False

        try:
            self.AMOUNT = float(amount)
            return True
        except ValueError:
            msg_amount = QMessageBox(QMessageBox.Warning, '错误', '请输入有效金额', QMessageBox.Ok, self)
            msg_amount.show()
            return False

    def inputItem(self):
        id = self.combox_id.currentText()   # str
        name = self.combox_name.currentText()   # str
        amount = self.AMOUNT  # str
        note = self.te_note.toPlainText()

        global member_df
        global item_df

        if self.valid_input(id, name, amount):
            # modify member sheet
            index_member = member_df.index[member_df['ID']==int(id)].tolist()[0]
            member_df.loc[index_member, '积分'] = self.AMOUNT
            member_df.to_excel('会员信息.xlsx', sheet_name='Sheet1', index=False)
            # modify item sheet
            new_info = pd.DataFrame({'ID':int(id), '消费金额':self.AMOUNT, '消费时间':datetime.datetime.now(), '备注':note}, index=[0])
            item_df_new = item_df.append(new_info, ignore_index=True)
            item_df_new.to_excel('消费条目.xlsx', sheet_name='Sheet1', index=False)
            item_df = item_df_new

            msg_confirm = QMessageBox(QMessageBox.Information, '确认', '已添加{}号会员{}的消费条目'.format(id,name), QMessageBox.Ok, self)
            msg_confirm.show()

            self.resetChoices()
            self.clear_input()


class PandasModel(QAbstractTableModel):
    def __init__(self, data, parent=None):
        super().__init__(parent=parent)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(),index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None


class ChkModWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("修改查询")
        self.setGeometry(300, 200, 900, 600)

        self.table = QTableView()
        self.model = PandasModel(member_df)
        self.table.setModel(self.model)
        self.table.resize(750, 600)
        self.table.setParent(self)
        self.table.setColumnWidth(2, 150)
        self.table.setColumnWidth(5, 190)

        btn_edit = QPushButton('修改信息', self)
        btn_edit.clicked.connect(self.myEdit)
        btn_edit.move(780, 50)
        btn_delete = QPushButton('删除会员', self)
        btn_delete.clicked.connect(self.myDelete)
        btn_delete.move(780, 90)
        btn_convert = QPushButton('转换积分', self)
        btn_convert.clicked.connect((self.myConvert))
        btn_convert.move(780, 130)
        self.show()

    def myEdit(self):
        index = (self.table.selectionModel().currentIndex())
        current_value = index.sibling(index.row(), index.column()).data()
        window_list.append(EditSubWindow(index.row(), index.column(), current_value, self))

    def myDelete(self):
        global member_df
        index = (self.table.selectionModel().currentIndex())
        id = member_df.iloc[index.row(), 0]
        name = member_df.iloc[index.row(), 1]
        result = QMessageBox.question(self, '提示', '确定删除{}号会员{}吗？'.format(id, name), QMessageBox.Yes | QMessageBox.No)
        if result == QMessageBox.Yes:
            member_df_dropped = member_df.drop(member_df[member_df['ID']==id].index).reset_index(drop=True)
            member_df_dropped.to_excel('会员信息.xlsx', sheet_name='Sheet1', index=False)
            member_df = member_df_dropped.copy()
            QMessageBox.information(self, 'INFO', '已删除', QMessageBox.Ok)
            self.model = PandasModel(member_df)
            self.table.setModel(self.model)
        else:
            pass

    def myConvert(self):
        global member_df
        index = (self.table.selectionModel().currentIndex())
        points = member_df.loc[index.row(), '积分']
        remaining = member_df.loc[index.row(), '余额']
        if points >= 100:
            member_df.loc[index.row(), '积分'] = points - 100
            member_df.loc[index.row(), '余额'] = remaining + 20
            member_df.to_excel('会员信息.xlsx', sheet_name='Sheet1', index=False)
            self.model = PandasModel(member_df)
            self.table.setModel(self.model)
            QMessageBox.information(self, 'INFO', '已转换积分', QMessageBox.Ok)
        else:
            QMessageBox.information(self, 'INFO', '积分未满100，不能转换', QMessageBox.Ok)


class EditSubWindow(QWidget):
    def __init__(self, row, col, current_value, chkmodwindow):
        super().__init__()
        self.row = row
        self.col = col
        self.changedValue = ''
        self.chkmodwindow = chkmodwindow
        self.initUI(current_value)

    def initUI(self, current_value):
        self.setWindowTitle('修改信息')
        self.resize(300, 200)

        lbl_orig = QLabel('现在值：', self)
        lbl_orig.move(70,60)
        lbl_orig_value = QLabel(str(current_value), self)
        lbl_orig_value.move(140,60)
        lbl_editto = QLabel('修改为：',self)
        lbl_editto.move(70, 80)
        self.le_edit = QLineEdit(self)
        self.le_edit.move(140, 80)
        self.le_edit.setFixedWidth(100)
        self.le_edit.textChanged[str].connect(self.setChange)

        btn_confirm = QPushButton('确定', self)
        btn_confirm.move(110, 140)
        btn_confirm.clicked.connect(self.saveChange)

        self.show()

    def valid_change(self, row, col, changedValue):
        header = member_df.columns.values[col]
        if header=='ID':
            msg_error = QMessageBox(QMessageBox.Warning, '错误', '不能修改ID', QMessageBox.Ok, self)
            msg_error.show()
            return False
        elif header=='电话' or header=='积分' or header=='余额':
            try:
                self.changedValue = int(changedValue)
            except ValueError:
                msg_error = QMessageBox(QMessageBox.Warning, '错误', '请输入有效数字', QMessageBox.Ok, self)
                msg_error.show()
                return False
        return True

    def setChange(self, s):
        self.changedValue = s

    def saveChange(self):
        global member_df
        if self.valid_change(self.row, self.col, self.changedValue):
            member_df.iloc[self.row, self.col] = self.changedValue
            member_df.to_excel('会员信息.xlsx', sheet_name='Sheet1', index=False)

            msg_confirm = QMessageBox(QMessageBox.Information, '确认', '已修改信息', QMessageBox.Ok, self)
            msg_confirm.show()

            self.chkmodwindow.model = PandasModel(member_df)
            self.chkmodwindow.table.setModel(self.chkmodwindow.model)


class ChkRecordWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.data = item_df
        self.id = None
        self.min_amount = None
        self.max_amount = None
        self.start_time = None
        self.end_time = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('查询记录')
        self.setGeometry(500, 300, 900, 600)

        self.table = QTableView()
        self.model = PandasModel(item_df)
        self.table.setModel(self.model)
        self.table.setParent(self)
        self.table.resize(900, 500)
        self.table.move(0, 100)
        self.table.setColumnWidth(0, 100)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(2, 250)
        self.table.setColumnWidth(3, 430)

        frame = QFrame(self)
        frame.resize(800, 100)

        lbl_id = QLabel('ID', self)
        lbl_id.move(20, 30)

        self.combox = QComboBox(self)
        self.combox.move(40, 30)
        self.combox.addItem('全部')
        id_str = [str(id) for id in set(item_df['ID'].values)]
        self.combox.addItems(id_str)
        self.combox.activated[str].connect(self.filterID)

        self.le_min = QLineEdit(self)
        self.le_min.setFixedWidth(60)
        self.le_min.move(150, 30)
        self.le_min.textChanged[str].connect(self.setMinAmount)
        lbl_min = QLabel('最小金额', self)
        lbl_min.setGeometry(150, 60, 70, 20)
        lbl_arrow1 = QLabel('--',self)
        lbl_arrow1.move(225, 35)
        self.le_max = QLineEdit(self)
        self.le_max.setFixedWidth(60)
        self.le_max.move(250, 30)
        self.le_max.textChanged[str].connect(self.setMaxAmount)
        lbl_max = QLabel('最大金额', self)
        lbl_max.setGeometry(250, 60, 70, 20)

        self.lbl_start = QLabel('起始时间', self)
        self.lbl_start.setFixedWidth(100)
        self.lbl_start.setAlignment(Qt.AlignRight)
        self.lbl_start.move(360, 30)
        lbl_arrow2 = QLabel('--', self)
        lbl_arrow2.move(480, 30)
        self.lbl_end = QLabel('结束时间', self)
        self.lbl_end.setFixedWidth(100)
        self.lbl_end.move(520, 30)

        self.btn_start = QPushButton('选择', self)
        self.btn_start.clicked.connect(self.filter_time_start)
        self.btn_start.setFixedWidth(40)
        self.btn_start.move(420, 55)
        self.btn_end= QPushButton('选择', self)
        self.btn_end.clicked.connect(self.filter_time_end)
        self.btn_end.setFixedWidth(40)
        self.btn_end.move(520, 55)

        lbl_sort = QLabel('排序', self)
        lbl_sort.setGeometry(650, 30, 50, 20)
        self.combox_sort = QComboBox(self)
        self.combox_sort.move(690, 30)
        self.combox_sort.addItem('请选择')
        self.combox_sort.addItems(['ID','金额正序','金额倒序','时间正序','时间倒序'])
        self.combox_sort.activated[str].connect(self.sort)


        self.btn_show = QPushButton('显示过滤', self)
        self.btn_show.setStyleSheet('background-color: #b7d7a8')
        self.btn_show.setGeometry(630, 70, 80, 25)
        self.btn_show.clicked.connect(self.show_filter)
        self.btn_reset = QPushButton('重置', self)
        self.btn_reset.setGeometry(720, 70, 50, 25)
        self.btn_reset.clicked.connect(self.reset_filter)

        self.show()

    def filterID(self, s):
        try:
            self.id = int(s)
        except ValueError:
            self.id = None

    def setMinAmount(self, s):
        self.min_amount = s

    def setMaxAmount(self, s):
        self.max_amount = s

    def filter_time_start(self):
        window_list.append(CalendarWindow('起始时间', self))

    def filter_time_end(self):
        window_list.append(CalendarWindow('结束时间', self))

    def sort(self, s):
        df = self.data.copy()
        if s == 'ID':
            df_sort = df.sort_values(by=['ID'])
        elif s == '金额正序':
            df_sort = df.sort_values(by=['消费金额'], ascending=True)
        elif s == '金额倒序':
            df_sort = df.sort_values(by=['消费金额'], ascending=False)
        elif s == '时间正序':
            df_sort = df.sort_values(by=['消费时间'], ascending=True)
        elif s == '时间倒序':
            df_sort = df.sort_values(by=['消费时间'], ascending=False)
        else:
            df_sort = df.copy()

        self.model = PandasModel(df_sort)
        self.table.setModel(self.model)


    def filter_valid(self):
        # check min and max
        if self.min_amount:
            try:
                self.min_amount = float(self.min_amount)
            except ValueError:
                QMessageBox.information(self, '错误', '请输入有效最小金额', QMessageBox.Ok)
                return False

        if self.max_amount:
            try:
                self.max_amount = float(self.max_amount)
            except ValueError:
                QMessageBox.information(self, '错误', '请输入有效最大金额', QMessageBox.Ok)
                return False

        if isinstance(self.min_amount, float) and isinstance(self.max_amount, float) and self.min_amount > self.max_amount:
            QMessageBox.information(self, '错误', '最小金额必须小于最大金额', QMessageBox.Ok)
            return False

        # check start and end date
        if isinstance(self.start_time, datetime.datetime) and isinstance(self.end_time, datetime.datetime):
            if self.start_time > self.end_time:
                QMessageBox.information(self, '错误', '起始结束时间有误', QMessageBox.Ok)
                return False

        return True


    def show_filter(self):
        if self.filter_valid():
            df_filter = item_df.copy()
            if self.id:
                df_filter = df_filter[df_filter['ID'] == self.id]
            if self.min_amount:
                df_filter = df_filter[df_filter['消费金额'] >= self.min_amount]
            if self.max_amount:
                df_filter = df_filter[df_filter['消费金额'] <= self.max_amount]
            if self.start_time:
                df_filter = df_filter[df_filter['消费时间'] >= self.start_time]
            if self.end_time:
                df_filter = df_filter[df_filter['消费时间'] <= self.end_time + datetime.timedelta(days=1)]

            self.data = df_filter
            self.combox_sort.setCurrentText('请选择')
            self.model = PandasModel(df_filter)
            self.table.setModel(self.model)

    def reset_filter(self):
        self.id = None
        self.min_amount = None
        self.max_amount = None
        self.start_time = None
        self.end_time = None

        self.combox.setCurrentText('全部')
        self.le_min.setText('')
        self.le_max.setText('')
        self.lbl_start.setText('起始时间')
        self.lbl_end.setText('结束时间')

        self.model = PandasModel(item_df)
        self.table.setModel(self.model)


class CalendarWindow(QWidget):
    def __init__(self, time_type, chkrecordwindow):
        super().__init__()
        self.type = time_type
        self.initUI()
        self.chkrecordwindow = chkrecordwindow

    def initUI(self, ):
        self.setWindowTitle('选择时间')
        self.resize(400, 300)

        cal = QCalendarWidget()
        cal.clicked[QDate].connect(self.set_time)
        lbl = QLabel('请选择'+self.type)
        lbl.setFont(QFont('',10))

        vbox = QVBoxLayout(self)
        vbox.addWidget(lbl)
        vbox.addWidget(cal)

        self.show()

    def set_time(self, date):
        date_str = date.toString('yyyy-MM-dd')
        if self.type == '起始时间':
            self.chkrecordwindow.lbl_start.setText(date_str)
            self.chkrecordwindow.start_time = datetime.datetime.strptime(date_str, '%Y-%m-%d')
        elif self.type == '结束时间':
            self.chkrecordwindow.lbl_end.setText(date_str)
            self.chkrecordwindow.end_time = datetime.datetime.strptime(date_str, '%Y-%m-%d')



if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont('微软雅黑', 11))
    app.setWindowIcon(QIcon('icon.jfif'))
    mainW = MainWindow()
    window_list = []
    member_df = pd.read_excel('会员信息.xlsx','Sheet1')
    item_df = pd.read_excel('消费条目.xlsx','Sheet1')
    app.exec_()