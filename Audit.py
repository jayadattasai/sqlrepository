from PyQt4 import QtCore, QtGui
global res
res=[]
global res1
res1=[]


import TestDataMiningTool
try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_MainWindowAudit(QtGui.QMainWindow):
    def __init__(self,queryDesc,parent=None):
        super(Ui_MainWindowAudit,self).__init__(parent)        
        self.queryDesc=queryDesc
        self.setupUi(self)
        
        
        
        
    def setupUi(self, MainWindowAudit):
        MainWindowAudit.setObjectName(_fromUtf8("MainWindowAudit"))
        MainWindowAudit.resize(796, 624)
        MainWindowAudit.setUnifiedTitleAndToolBarOnMac(False)
        self.centralwidget = QtGui.QWidget(MainWindowAudit)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.textBrowser_bv = QtGui.QTextBrowser(self.centralwidget)
        self.textBrowser_bv.setGeometry(QtCore.QRect(20, 198, 361, 400))
        self.textBrowser_bv.setObjectName(_fromUtf8("textBrowser_bv"))
        self.textBrowser_av = QtGui.QTextBrowser(self.centralwidget)
        self.textBrowser_av.setGeometry(QtCore.QRect(409, 198, 361, 400))
        self.textBrowser_av.setObjectName(_fromUtf8("textBrowser_av"))
        self.tableView = QtGui.QTableView(self.centralwidget)
        self.tableView.setGeometry(QtCore.QRect(19, 31, 755, 141))
        self.tableView.setObjectName(_fromUtf8("tableView"))
        self.model = QtGui.QStandardItemModel()
        self.tableView.setModel(self.model)
        curdesc= ['Module','Query Description','Updated On','By User','Type']
        self.model.setHorizontalHeaderLabels(curdesc)
##        self.tableView.setColumnWidth(0,100)
##        self.tableView.setColumnWidth(1,350)
##        self.tableView.setColumnWidth(2,120)
##        self.tableView.setColumnWidth(3,75)
##        self.tableView.setColumnWidth(4,75)
        self.tableView.setFrameShape(QtGui.QFrame.NoFrame)
        self.tableView.setFrameShadow(QtGui.QFrame.Sunken)
        self.tableView.setMidLineWidth(0)
        self.tableView.setEditTriggers(QtGui.QAbstractItemView.SelectedClicked)
        self.tableView.setProperty("showDropIndicator", True)
        self.tableView.setDragEnabled(False)
        self.tableView.setDefaultDropAction(QtCore.Qt.CopyAction)
        self.tableView.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
        
        self.tableView.setAlternatingRowColors(False)
        self.tableView.setSelectionMode(QtGui.QAbstractItemView.ExtendedSelection)
        self.tableView.setTextElideMode(QtCore.Qt.ElideNone)
        self.tableView.setGridStyle(QtCore.Qt.SolidLine)
        self.tableView.setWordWrap(True)
        self.tableView.setCornerButtonEnabled(True)
        self.tableView.setObjectName(_fromUtf8("tableView"))        
        self.tableView.horizontalHeader().setCascadingSectionResizes(False)
        self.tableView.horizontalHeader().setResizeMode(QtGui.QHeaderView.Stretch)
        self.tableView.horizontalHeader().setStretchLastSection(False)
        self.tableView.verticalHeader().setDefaultSectionSize(20)
        
        self.tableView.clicked.connect(self.showVal)

        
        self.label = QtGui.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(170, 178, 81, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName(_fromUtf8("label"))
        self.label_2 = QtGui.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(550, 178, 81, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName(_fromUtf8("label_2"))
        self.label_3 = QtGui.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(32, 10, 261, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setObjectName(_fromUtf8("label_3"))
##        self.pushButton = QtGui.QPushButton(self.centralwidget)
##        self.pushButton.setGeometry(QtCore.QRect(690, 580, 75, 23))
##        self.pushButton.setObjectName(_fromUtf8("pushButton"))
##        self.pushButton.clicked.connect(self.close) ####################
        MainWindowAudit.setCentralWidget(self.centralwidget)
        self.statusbar = QtGui.QStatusBar(MainWindowAudit)
        self.statusbar.setObjectName(_fromUtf8("statusbar"))
        MainWindowAudit.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindowAudit)
        QtCore.QMetaObject.connectSlotsByName(MainWindowAudit)
        self.populate(self.queryDesc)

    def retranslateUi(self, MainWindowAudit):
        MainWindowAudit.setWindowTitle(_translate("MainWindowAudit", "Audit Information - Test Data Mining Tool", None))
        self.label.setText(_translate("MainWindowAudit", "Before Value", None))
        self.label_2.setText(_translate("MainWindowAudit", "After Value", None))
        self.label_3.setText(_translate("MainWindowAudit", "Audit History:", None))
##        self.pushButton.setText(_translate("MainWindowAudit", "Close", None))

    def showVal(self):
        sel =self.tableView.selectionModel().selectedRows()         
        
        
        if  sel[0].row()==0:
            self.textBrowser_bv.setText("NULL")
            if not res:
                self.textBrowser_av.setText(res1[0][5])
            else:
                self.textBrowser_av.setText(res[0][5])
        else:
            self.textBrowser_bv.setText(res[sel[0].row()-1][5])
            self.textBrowser_av.setText(res[sel[0].row()-1][6])
        
        
        bb=self.textBrowser_bv.toPlainText().split(' ')
        aa=self.textBrowser_av.toPlainText().split(' ')
        cc= set(bb)
        dd=[x for x in aa if x not in cc]
        
        findd= ' '.join(dd)
        colored ="<font color=red>"+findd+"</font>"
        w=self.textBrowser_av.toPlainText().replace(findd,colored)
        self.textBrowser_bv.setHtml("<pre>"+self.textBrowser_bv.toPlainText()+"</pre>")
        self.textBrowser_av.setHtml("<pre>"+w+"</pre>")
        QtGui.qApp.processEvents()

    def populate(self,queryDesc):
        import pyodbc
        import datetime
        MDB =TestDataMiningTool.URL; DRV='{Microsoft Access Driver (*.mdb)}'
        con = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV,MDB))
        cur = con.cursor()
##        params =
        SQL1="SELECT Module,Description, [Created date], UserId, 'Insert',Query FROM Table1 where Table1.[Description]=?"
##        SQL1="SELECT * FROM Table1 where Table1.[Description]=?" datetime.datetime.strftime('%d %b %Y')
        rows=cur.execute(SQL1,queryDesc).fetchall()
        global res1
        res1=rows
        
        
        if (len(rows)>0) :
            g= []
            for i in  range(len(rows)):
                for j in range(len(cur.description)-1):
                    rrow =QtGui.QStandardItem(str(rows[i][j]))
                    g.append(rrow)                    
                self.model.appendRow(g)
                g=[]
        SQL2="SELECT Module,QueryDesc,Editdate, UserId, Type, [BeforeVal],[AfterVal] FROM Audit where Audit.[QueryDesc]=?"
        rows=cur.execute(SQL2,queryDesc).fetchall()
        global res
        res =rows
        
        if (len(rows)>0) :
            g= []
            for i in  range(len(rows)):
                for j in range(len(cur.description)-2):
                    rrow =QtGui.QStandardItem(str(rows[i][j]))
                    g.append(rrow)                    
                self.model.appendRow(g)
                g=[]       
                
        cur.close()
        con.close()

if __name__ == "__main__":
    import sys
    app = QtGui.QApplication(sys.argv)
    MainWindowAudit = QtGui.QMainWindow()
    ui = Ui_MainWindowAudit()
    ui.setupUi(MainWindowAudit)
    MainWindowAudit.show()
    sys.exit(app.exec_())

