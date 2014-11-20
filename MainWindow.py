#!/usr/bin/env python
#coding=utf-8
 
import sys
import os
from PyQt4.QtCore import *
from PyQt4.QtGui import * 

from sqlitedb import DB
import scheme
from openpyxl import load_workbook
from excel import *

reload(sys)
sys.setdefaultencoding('utf8')                
 
class MainWindow(QMainWindow):
	def __init__(self,parent=None):
		super(MainWindow,self).__init__(parent)

		fileNewAction=QAction(u"导入",self)
		fileNewAction.setShortcut(QKeySequence.New)

		fileDumpAction=QAction(u"导出",self)
		# helpText = "Create a new file"
		# fileNewAction.setToolTip(helpText)
		# fileNewAction.setStatusTip(helpText)
		self.connect(fileNewAction,SIGNAL("triggered()"),self.fileNew)
		
		self.fileMenu = self.menuBar().addMenu("&File")
		self.fileMenu.addAction(fileNewAction)
		# self.fileMenu.addAction(fileDumpAction)
		
		filetoolbar = self.addToolBar("File")
		filetoolbar.addAction(fileNewAction)
		# filetoolbar.addAction(fileDumpAction)
		    
		self.status = self.statusBar()
		self.status.showMessage("This is StatusBar",5000)
		self.setWindowTitle(u"预算分析")

		self.textedit = QTextEdit()
		self.textedit.setText("This is a TextEdit!")

		self.tab_widget = QTabWidget(self)
		self.tab_widget.setTabsClosable(True)
		self.tab_widget.tabBar().setMovable(True)
		self.tab_widget.tabCloseRequested.connect(self.close_handler)
      
		self.listwidget = DataList(self.tab_widget)
		# for i in range(1, 50):
		# 	self.listwidget.addItem("table%r" % i)
		

		lefttab = QTabWidget(self)
		lefttab.addTab(self.listwidget, u"数据")

		self.querypage = QueryPage(self.listwidget)
      
		self.treewidget = QTreeWidget()
		self.treewidget.setHeaderLabels(['This','is','a','TreeWidgets!'])
      
		splitter = QSplitter(self)

		# splitter.addWidget(self.textedit)
		splitter.addWidget(lefttab)
		splitter.addWidget(self.tab_widget)
		splitter.addWidget(self.querypage)


		# splitter.addWidget(self.tab_widget)
		# splitter.addWidget(self.tab_widget)
		# splitter.setOrientation(Qt.Vertical)
		self.setCentralWidget(splitter)
		# self.setCentralWidget(self.tab_widget)
		splitter.setSizes([100, 400, 200])
		self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)  
		self.activateWindow()

		self.resize(800,600)

	def fileNew(self):
		self.status.showMessage("You have created a new file!",9000)
		fileName = QFileDialog.getOpenFileName(
			self,
			self.tr("Open DOC"),
			"./",
			self.tr("Doc Files(*.xls *.xlsx)")
		)
		print fileName
		# self.slotStart()
		if fileName:
			self.load_file(str(fileName).decode("utf-8"))
		self.listwidget.refresh_table()
		self.querypage.filenames()

		# with open(filename, 'r') as f:
		# filedata = f.read()
		# text_widget = QTextEdit(self.tab_widget)
		# # text_widget.setText(filedata)
		# self.status.showMessage(fileName,9000)
		# item = QListWidgetItem("table%s" % fileName)
		# self.listwidget.addItem(item)
		# self.listwidget.clicked.connect(self.showDialog)
		# self.tab_widget.addTab(text_widget, os.path.basename(filename))
	def showDialog(self, item):
		print "click"
		print item, item.text()

	def close_handler(self, index):
		print "close_handler called, index = %s" % index
		self.tab_widget.removeTab(index)

	def slotStart(self):  
		num=10  

		progressDialog=QProgressDialog(self)  
		progressDialog.setWindowModality(Qt.WindowModal)  
		progressDialog.setMinimumDuration(5)  
		progressDialog.setWindowTitle(self.tr("请等待"))  
		progressDialog.setLabelText(u"拷贝...")  
		# progressDialog.setCancelButtonText(self.tr("cancel"))  
		progressDialog.setRange(0,num)  

		for i in range(num):  
		    progressDialog.setValue(i)  
		    QThread.msleep(100)  
		    if progressDialog.wasCanceled():  
		        return

	def load_file(self, filename=''):
		if not filename:
			return false
		print filename
		conn = DB("haya.db3")

		

		sheets = scheme.monthList

		progressDialog=QProgressDialog(self)  
		progressDialog.setWindowModality(Qt.WindowModal)  
		progressDialog.setMinimumDuration(5)  
		progressDialog.setWindowTitle(u"waiting....")  
		progressDialog.setLabelText(u"import...")  
		# progressDialog.setCancelButtonText(self.tr("cancel"))  
		progressDialog.setRange(0, len(sheets))
		
		progressDialog.show()

		wb = load_workbook(filename = filename)
		i = 1
		ufields = scheme.fields[1:]			
		for x in sheets:
			print x

			ws = wb.get_sheet_by_name(name = x)
			if not ws:
				continue
			print "Work Sheet Titile:",ws.title 
			rows = ws.rows
			for row in rows[7:]:
				if not row[1].value:
					break
				data = [str(d.value) if d.value else "0" for d in list(row[1:len(ufields)+1])]
				sql = "insert into haya_budget values (null, '%s', '%s', '%s')" % (os.path.basename(filename), x, "','".join(data))
				print sql
				conn.runSql(sql)

			progressDialog.setValue(i)  
			i = i+1

		
		

		ws = wb.get_sheet_by_name(name=u'社保住房合规增项') #社保住房合规增项
		
		if ws:
			print ws.title
			rows = ws.rows
			for row in rows[7:]:
				if not row[1].value:
					break
				data = [
					str(row[1].value) if row[1].value else "0", 
					str(row[2].value) if row[2].value else "0",
					str(row[3].value) if row[3].value else "0",
					str(row[10].value) if row[10].value else "0",
					str(row[11].value) if row[11].value else "0",
					str(row[12].value) if row[12].value else "0",
					str(row[20].value) if row[20].value else "0",
					str(row[21].value) if row[21].value else "0",
					str(row[22].value) if row[22].value else "0",
					str(row[23].value) if row[23].value else "0",
					str(row[24].value) if row[24].value else "0",
					str(row[25].value) if row[25].value else "0",
					str(row[26].value) if row[26].value else "0",
					str(row[27].value) if row[27].value else "0",
					str(row[28].value) if row[28].value else "0",
					str(row[29].value) if row[20].value else "0",
					str(row[30].value) if row[30].value else "0",
					str(row[31].value) if row[31].value else "0",
					str(row[32].value) if row[32].value else "0",
				]
				sql = """insert into haya_addons (
					id,
					[月份],
					[公司名称],
					[工资发放地],
					[实际工作地],
					[状态],
					[一级部门],
					[二级部门],
					[月社保基数增加额],
					[社保基数合规月数],
					[月养老增加额],
					[月工伤增加额],
					[月失业增加额],
					[月生育增加额],
					[月医疗增加额],
					[社保合规总额],
					[住房基数增加额],
					[合规月数],
					[月住房公积金增加额],
					[住房合规总额],
					[社保稽查补缴]
				)
				 values (null, '%s', '%s')""" % (os.path.basename(filename), "','".join(data))
				print sql
				conn.runSql(sql)
			

		conn.closeDb()
		progressDialog.close()

class DataList(QListWidget):
    def __init__(self, tab):
        QListWidget.__init__(self)
        self.init_table()
        self.itemDoubleClicked.connect(self.item_click)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.tab_widget = tab

    def init_table(self):
        # for item_text in ['item1', 'item2', 'item3']:
        #     item = QListWidgetItem(item_text)
        #     self.addItem(item)
		# sql = "SELECT name FROM sqlite_master WHERE type='table'"
		sql = "SELECT distinct 公司名称 FROM haya_budget"
		print sql
		conn = DB("haya.db3")
		rs = conn.getData(sql)
		conn.closeDb()
		print rs
		for tbname in rs:
			item = QListWidgetItem(tbname[0])
			item.setData(1,"aaa")
			self.addItem(item)
    def refresh_table(self):
		self.clear()
		self.init_table()

    def item_click(self, item):
		print item, item.text()
		print item.data(1)
		# self.table = DataTable()
		tbcount = self.tab_widget.count()
		for i in range(tbcount):
			tname = self.tab_widget.tabText(i)
			if tname == item.text():
				self.tab_widget.setCurrentIndex(i)
				return
		tab1 = DataTable(item.text())

		# tab2 = self.treewidget
		print "index of", self.tab_widget.indexOf(tab1)
		self.tab_widget.addTab(tab1, item.text()) 
		print self.tab_widget.count()

class DataTable(QWidget):
	def __init__(self, name):
		QWidget.__init__(self)
		self.conn = DB("haya.db3")
		self.createTable(name)
		self.conn.closeDb()
    
	@pyqtSlot(int)    
	def on_view_horizontalHeader_sectionClicked(self, logicalIndex):
		self.logicalIndex   = logicalIndex
		self.menuValues     = QMenu(self)
		self.signalMapper   = QSignalMapper(self)

		self.comboBox.blockSignals(True)
		self.comboBox.setCurrentIndex(self.logicalIndex)
		self.comboBox.blockSignals(True)

		valuesUnique = [    self.model.item(row, self.logicalIndex).text()
                            for row in range(self.model.rowCount())
                            ]

		actionAll = QAction("All", self)
		actionAll.triggered.connect(self.on_actionAll_triggered)
		self.menuValues.addAction(actionAll)
		self.menuValues.addSeparator()

		for actionNumber, actionName in enumerate(sorted(list(set(valuesUnique)))):              
			action = QAction(actionName, self)
			self.signalMapper.setMapping(action, actionNumber)  
			action.triggered.connect(self.signalMapper.map)  
			self.menuValues.addAction(action)

		self.signalMapper.mapped.connect(self.on_signalMapper_mapped)  

		headerPos = self.view.mapToGlobal(self.horizontalHeader.pos())        
		# print "headerPos", headerPos, self.horizontalHeader.height(), self.logicalIndex
		posY = headerPos.y() + self.horizontalHeader.height()
		posX = headerPos.x() + self.horizontalHeader.sectionPosition(self.logicalIndex)

		self.menuValues.exec_(QPoint(posX, posY))

	@pyqtSlot()
	def on_actionAll_triggered(self):
		filterColumn = self.logicalIndex
		filterString = QRegExp(  "",
                                        Qt.CaseInsensitive,
                                        QRegExp.RegExp
                                        )

		self.proxy.setFilterRegExp(filterString)
		self.proxy.setFilterKeyColumn(filterColumn)

	@pyqtSlot(int)
	def on_signalMapper_mapped(self, i):
		stringAction = self.signalMapper.mapping(i).text()
		filterColumn = self.logicalIndex
		filterString = QRegExp(  stringAction,
		                                Qt.CaseSensitive,
		                                QRegExp.FixedString
		                                )

		self.proxy.setFilterRegExp(filterString)
		self.proxy.setFilterKeyColumn(filterColumn)


	@pyqtSlot(str)
	def on_lineEdit_textChanged(self, text):
		search = QRegExp(    text,
                                    Qt.CaseInsensitive,
                                    QRegExp.RegExp
                                    )

		self.proxy.setFilterRegExp(search)

	@pyqtSlot(int)
	def on_comboBox_currentIndexChanged(self, index):
		self.proxy.setFilterKeyColumn(index)

	def createTable(self, name):
		self.view           = QTableView()
		self.comboBox       = QComboBox()
		self.model = QStandardItemModel(self)
		self.proxy = QSortFilterProxyModel(self)
		self.model.setHorizontalHeaderLabels(scheme.fields)
		sql = "select count(*) as count from haya_budget where 公司名称='%s'" % name
# 		print sql
		rs = self.conn.getLine(sql)
		rowcount = rs[0]
		colcount = len(scheme.fields)
		print rowcount, colcount
		sql = "select * from haya_budget where 公司名称='%s'" % name
		rs = self.conn.getData(sql)

		for i in range(len(rs)):
			data = list(rs[i])
			del data[0]
			del data[0]
			self.model.invisibleRootItem().appendRow(
                [   QStandardItem(str(column).decode('utf-8'))    
                    for column in data
                    ]
                )
		self.proxy.setSourceModel(self.model)
		self.view.setModel(self.proxy)
		self.horizontalHeader = self.view.horizontalHeader()
		self.horizontalHeader.sectionClicked.connect(self.on_view_horizontalHeader_sectionClicked)
		
		mainLayout = QVBoxLayout()
		mainLayout.addWidget(self.view)
        
		self.setLayout(mainLayout)



# class DataTable(QTableWidget):
# 	def __init__(self, name):
# 		QTableWidget.__init__(self)
# 		self.conn = DB("haya.db3")
# 		self.createTable(name)
		

# 	def createTable(self, name=""):
# 		# self.table = QTableWidget(300,7)
# 		sql = "select count(*) as count from haya_budget where 公司名称='%s'" % name
# 		print sql
# 		rs = self.conn.getLine(sql)
# 		self.setRowCount(rs[0])
# 		self.setColumnCount(len(scheme.fields))
# 		# self.setHorizontalHeaderLabels([u'姓名','MON','TUE','WED',
#   #                                             'THU','FIR','SAT'])
# 		self.setHorizontalHeaderLabels(scheme.fields)
# 		self.setAlternatingRowColors(True)
# 		self.setEditTriggers(QTableWidget.DoubleClicked)
# 		self.setSortingEnabled(True)

# 		sql = "select * from haya_budget where 公司名称='%s'" % name
# 		rs = self.conn.getData(sql)
# 		for i in range(len(rs)):
# 			data = list(rs[i])
# 			del data[0]
# 			del data[0]
# 			for j in range(len(data)):
# 				print data[j]
# 				newItem = QTableWidgetItem(str(data[j]).decode('utf-8'))
# 				self.setItem(i, j, newItem)




# class FilePropertiesDlg(QfileDialog):
 
# 	def __init__(self,parent=None):
# 		super(FilePropertiesDlg, self).__init__(self,Qwidget parent=None,Qstring caption=Qstring(),Qstring directory =Qstring(),Qstring filter=Qstring())

# 	def openFile(self):
# 		fileName = QFileDialog.getOpenFileName(self,self.tr(“Open Image”),Qstring(),self.tr("Image Files(*.png *.jpg *.bmp)"))

class DataFilter(QWidget):
    def __init__(self,parent=None):
        QWidget.__init__(self,parent)
 
        self.setWindowTitle('grid layout2')
 
        title = QLabel('Tltle')
        author = QLabel('Author')
        review = QLabel('Review')
 
        titleEdit = QDateTimeEdit(QDate.currentDate())
        authorEdit = QLineEdit()
        reviewEdit = QTextEdit()
 
        grid = QGridLayout()
        grid.setSpacing(10)
 
        grid.addWidget(title,1,0)
        grid.addWidget(titleEdit,1,1)
 
        grid.addWidget(author,2,0)
        grid.addWidget(authorEdit,2,1)
 
        grid.addWidget(review,3,0)
        grid.addWidget(reviewEdit,3,1,5,1)
 
        self.setLayout(grid)
        # self.resize(350,300)
class QueryPage(QWidget):
    def __init__(self, listwidget=None):
        super(QueryPage, self).__init__()
        self.listwidget = listwidget
        packagesGroup = QGroupBox(u"条件")
  
        nameLabel = QLabel(u"部门")
        nameEdit = QLineEdit()
  
        dateLabel = QLabel(u"月份:")
        self.dateEdit = QComboBox()
        self.dateEdit.addItem(u"全部", 0)
        self.dateEdit.addItems(scheme.monthList)
        # self.connect(dateEdit,SIGNAL("activated(const QString &text)"),self.doDump)
        self.dateEdit.currentIndexChanged['QString'].connect(self.doDump)
        # dateEdit.valueChanged.connect(self.doDump)

        levelLabel = QLabel(u"条件")
        # levelOptions = QComboBox()
        # levelOptions.addItem(u"全部")
        # levelOptions.addItems(scheme.fields)
        levelOptions = QListWidget()
        levelOptions.addItem(u"全部")
        levelOptions.addItems(scheme.fields[1:20])
        levelOptions.setSelectionMode(QAbstractItemView.ExtendedSelection)

  
        # releasesCheckBox = QCheckBox("Releases")
        # upgradesCheckBox = QCheckBox("Upgrades")
  
        # hitsSpinBox = QSpinBox()
        # hitsSpinBox.setPrefix("Return up to ")
        # hitsSpinBox.setSuffix(" results")
        # hitsSpinBox.setSpecialValueText("Return only the first result")
        # hitsSpinBox.setMinimum(1)
        # hitsSpinBox.setMaximum(100)
        # hitsSpinBox.setSingleStep(10)

        # optionMenu = QComboBox()
        # optionMenu.addItems([u"宋体", u"黑体", u"仿宋",
        #                              u"隶书", u"楷体"])
  
        # startQueryButton = QPushButton(u"筛选")
        dumpButton = QPushButton(u"导出")
        self.connect(dumpButton,SIGNAL("clicked()"),self.doDump)

        removeButton = QPushButton(u"清除")
        self.removeMenu = QComboBox()
        self.filenames()
        self.connect(removeButton,SIGNAL("clicked()"),self.clearFileList)

        packagesLayout = QGridLayout()
        # packagesLayout.addWidget(nameLabel, 0, 0)
        # packagesLayout.addWidget(optionMenu, 0, 1)
        packagesLayout.addWidget(dateLabel, 1, 0)
        packagesLayout.addWidget(self.dateEdit, 1, 1)
        packagesLayout.addWidget(levelLabel, 2, 0)
        packagesLayout.addWidget(levelOptions, 2, 1)
        # packagesLayout.addWidget(releasesCheckBox, 2, 0)
        # packagesLayout.addWidget(upgradesCheckBox, 3, 0)
        # packagesLayout.addWidget(hitsSpinBox, 4, 0, 1, 2)
        packagesGroup.setLayout(packagesLayout)
  
        mainLayout = QVBoxLayout()
        # mainLayout.addWidget(packagesGroup)
        mainLayout.addSpacing(12)
        # mainLayout.addWidget(startQueryButton)
        mainLayout.addWidget(dumpButton)
        mainLayout.addStretch(1)
        mainLayout.addWidget(self.removeMenu)
        mainLayout.addWidget(removeButton)
  
        self.setLayout(mainLayout)

    def doDump(self, text=''):
		print "dumping", text;
		# print self.dateEdit.itemData(self.dateEdit.currentIndex()).toPyObject()
		# print self.dateEdit.currentText()
		progressDialog=QProgressDialog(self)  
		progressDialog.setWindowModality(Qt.WindowModal)  
		progressDialog.setMinimumDuration(5)  
		progressDialog.setWindowTitle(u"waiting....")  
		progressDialog.setLabelText(u"正在导出...")  
		# progressDialog.setCancelButtonText(self.tr("cancel"))  
		progressDialog.setRange(0, 2)

		progressDialog.show()
		progressDialog.setValue(1) 
		dumpfile()
		progressDialog.close()
		QMessageBox.warning(self,u"文件导出",u"导出成功")

    def filenames(self):
    	self.removeMenu.clear()
    	sql = "SELECT distinct 文件 FROM haya_budget"
        conn = DB("haya.db3")
        rs = conn.getData(sql)
        conn.closeDb()
        for f in rs:
        	self.removeMenu.addItem(f[0])

    def clearFileList(self):
    	filename = self.removeMenu.currentText()
    	button=QMessageBox.question(self,"Question",  
                                    u"清空来自%s的数据" % filename,  
                                    QMessageBox.Ok|QMessageBox.Cancel,  
                                    QMessageBox.Cancel) 
    	if button==QMessageBox.Ok:  
			sql = "delete FROM haya_budget where 文件='%s'" % filename
			sql1 = "delete FROM haya_addons where 月份='%s'" % filename
			conn = DB("haya.db3")
			conn.runSql(sql)
			conn.runSql(sql1)
			conn.closeDb()
			self.filenames()
			self.listwidget.refresh_table()
    	else:
			return

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("PyQt MainWindow")
    # app.setWindowIcon(QIcon("./images/icon.png"))
    form = MainWindow()
    form.show()
    sys.exit(app.exec_())
 
main()