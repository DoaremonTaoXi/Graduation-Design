import sys
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QMessageBox, QDialog, \
QGroupBox, QLabel, QLineEdit, QHBoxLayout, QTableWidgetItem, QMainWindow
from PyQt6.QtGui import QIcon, QPixmap, QImage
from PyQt6.QtCore import Qt, pyqtSignal, QObject
from PyQt6 import uic, QtWidgets
from threading import  Thread
import xlwings as xw
import time
import myFct

# 信号库
class SignalStore(QObject):
    # 定义信号
    signal_warning = pyqtSignal()
    signal_updateTableSheets = pyqtSignal(tuple)
    signal_updateTableResults = pyqtSignal(tuple)
    signal_updateProgressBar = pyqtSignal(int)
    signal_initProgressBar = pyqtSignal(int)
    signal_finished = pyqtSignal(float)

# 实例化
so = SignalStore()

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        # Load the UI file
        self.ui = uic.loadUi("demo.ui")

        # 获取QTableWidget实例
        self.TableResults = self.ui.findChild(QtWidgets.QTableWidget, "TableResults")
        self.TableSheets = self.ui.findChild(QtWidgets.QTableWidget, "TableSheets")

        # 将进度条拉回至原点
        self.ui.progressBar.reset()

        # Connect signals and slots
        self.ui.Button_Import.clicked.connect(self.on_Button_Import_clicked)
        self.ui.Button_Run.clicked.connect(self.on_Button_Run_clicked)
        self.ui.Button_Cancel.clicked.connect(QApplication.instance().quit)

        so.signal_warning.connect(self.signal_warning_slot)
        so.signal_updateTableSheets.connect(self.signal_updateTableSheets_slot)
        so.signal_updateTableResults.connect(self.signal_updateTableResults_slot)
        so.signal_updateProgressBar.connect(self.signal_updateProgressBar_slot)
        so.signal_initProgressBar.connect(self.signal_initProgressBar_slot)
        so.signal_finished.connect(self.signal_finished_slot)
    
        # 统计进行中标记，不能同时做两个检查
        self.ongoing = False

    # Signal的Slot函数
    def signal_warning_slot(self):
        QMessageBox.warning(self.ui,'Warning','Please import an Excel file first.')

    def signal_updateTableSheets_slot(self, tuple):
        # TableResults初始化
        self.init_TableSheets()
        # 输出结果
        sheetnames, sheetNrows, sheetNcols = tuple
        for i in range(len(sheetnames)):
            self.TableSheets.insertRow(i)
            item0 = QTableWidgetItem(sheetnames[i]); item0.setTextAlignment(Qt.AlignmentFlag.AlignVCenter)  # 设置文本垂直居中
            self.TableSheets.setItem(i, 0, item0)
            item1 = QTableWidgetItem(sheetNrows[i]); item1.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # 设置文本水平和垂直居中
            self.TableSheets.setItem(i, 1, item1)
            item2 = QTableWidgetItem(sheetNcols[i]); item2.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # 设置文本水平和垂直居中
            self.TableSheets.setItem(i, 2, item2)

    def signal_updateTableResults_slot(self,tuple):
        # TableResults初始化
        self.init_TableResults()
        # 输出结果
        Error_flag, Error_sheet, Error_cell, Error_value = tuple
        if Error_flag:
            for i in range(len(Error_cell)):
                self.TableResults.insertRow(i)
                item0 = QTableWidgetItem(Error_sheet[i]); item0.setTextAlignment(Qt.AlignmentFlag.AlignVCenter)  # 设置文本垂直居中
                self.TableResults.setItem(i, 0, item0)
                item1 = QTableWidgetItem(Error_cell[i]); item1.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # 设置文本水平和垂直居中
                self.TableResults.setItem(i, 1, item1)
                item2 = QTableWidgetItem(Error_value[i]); item2.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # 设置文本水平和垂直居中
                self.TableResults.setItem(i, 2, item2)

    def signal_updateProgressBar_slot(self, value):
        self.ui.progressBar.setValue(value)

    def signal_initProgressBar_slot(self, value):
        self.ui.progressBar.setRange(0, value)

    def signal_finished_slot(self, value):
        QMessageBox.information(self.ui,'Run successfully',f'elapsed time:  {value:.2f}  s')
    
    # Button的slot函数
    def on_Button_Import_clicked(self):
         # Show a file dialog to select one or more Excel files
        filename, _ = QFileDialog.getOpenFileName(self, "Import Excel Files", "", "Excel Files (*.xls *.xlsx)")
        self.filename = filename
        if filename:
            # Set the text of the label to the imported file names
            self.ui.lineEdit_filename.setText(filename)

    def on_Button_Run_clicked(self):
        def workerThreadFunc():
            self.ongoing = True

            if hasattr(self, 'filename'):  # 判断self里有无filename
                start_time = time.time()  # 记录开始时间

                # 记录信息初始化
                sheetnames = []; sheetNrows = []; sheetNcols = []
                Error_flag = 0; Error_sheet = []; Error_cell = []; Error_value = []

                with xw.App(visible = False, add_book = False) as app:  # 启动Excel程序
                    workbook = app.books.open(self.filename)
                    worksheets = workbook.sheets  # 获取工作簿中所有的工作表

                    so.signal_initProgressBar.emit(worksheets.count)  # 设置进度条范围

                    for worksheet in worksheets:
                        # 记录表格信息（名称、行数、列数）
                        sheetnames.append(worksheet.name)
                        rng = worksheet.used_range
                        sheetNrows.append(str(rng.rows.count))
                        sheetNcols.append(str(rng.columns.count))
                        
                        # 运行检查程序，纪录错误信息
                        ErrorName_flag, ErrorName_cell, ErrorName_value = myFct.CheckName(worksheet)
                        ErrorNumber_flag, ErrorNumber_cell, ErrorNumber_value = myFct.CheckNumber(worksheet)
                        ErrorContent_flag, ErrorContent_cell, ErrorContent_value = myFct.CheckContent(worksheet)
                        # 将错误信息进行整合
                        Error_flag = ErrorName_flag or ErrorNumber_flag or ErrorContent_flag or Error_flag
                        Error_cell.extend(ErrorName_cell + ErrorNumber_cell + ErrorContent_cell)
                        Error_value.extend(ErrorName_value + ErrorNumber_value + ErrorContent_value)
                        for i in range(len(ErrorName_cell + ErrorNumber_cell + ErrorContent_cell)):
                            Error_sheet.append(worksheet.name)

                        so.signal_updateProgressBar.emit(worksheet.index)
                        # print(f'{worksheet.name} Run successfully')
                
                # 输出检查结果
                so.signal_updateTableSheets.emit((sheetnames, sheetNrows, sheetNcols))
                so.signal_updateTableResults.emit((Error_flag, Error_sheet, Error_cell, Error_value))         

                end_time = time.time()
                elapsed_time = end_time - start_time
                so.signal_finished.emit(elapsed_time)

            else:
                # 发出信息，通知主线程进行进度处理
                so.signal_warning.emit()
                return

            self.ongoing = False

        if self.ongoing:
            QMessageBox.warning(self.ui,'Warning','Program is ongoing, Please wait a monment')
            return
        
        # 创建新线程
        worker = Thread(target=workerThreadFunc)
        worker.start()

    # TableResults初始化
    def init_TableResults(self):
        # 清除之前内容
        self.TableResults.setRowCount(0)
        # 设置表格的列数
        self.TableResults.setColumnCount(3)
        #设置表格的水平表头
        headers = ['SHEETNAME', 'CELL', 'VALUE']
        self.TableResults.setHorizontalHeaderLabels(headers)
        # 设定第1列的宽度为 180像素
        self.TableResults.setColumnWidth(0, 180)
        # 设定第2列的宽度为 100像素
        self.TableResults.setColumnWidth(1, 100)
        # 设定第3列的宽度为 180像素
        self.TableResults.setColumnWidth(2, 180)

    # TableSheets初始化
    def init_TableSheets(self):
        # 清除之前内容
        self.TableSheets.setRowCount(0)
        # 设置表格的列数
        self.TableSheets.setColumnCount(3)
        #设置表格的水平表头
        headers = ['SHEETNAME', 'Nrows', 'Ncols']
        self.TableSheets.setHorizontalHeaderLabels(headers)
        # 设定第1列的宽度为 100像素
        self.TableSheets.setColumnWidth(0, 180)
        # 设定第2列的宽度为 80像素
        self.TableSheets.setColumnWidth(1, 80)
        # 设定第3列的宽度为 80像素
        self.TableSheets.setColumnWidth(2, 80)



# 主函数
if __name__ == '__main__':
    # QApplication 提供了整个图形界面程序的底层管理功能
    app = QApplication([])
    # 创建Win，引用文件中的Ui_MainWindow类
    MyWin = MainWindow()
    # 显示界面
    MyWin.ui.show()
    # 进入QApplication的事件处理循环，接收用户的输入事件
    sys.exit(app.exec())