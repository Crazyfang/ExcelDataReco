import sys
import Surface
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtCore import QBasicTimer, QStringListModel, QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
from Main import ProcessData
import os


class ThreadTransfer(QThread):
    signOut = pyqtSignal(str, float)

    def __init__(self, path):
        super(ThreadTransfer, self).__init__()
        self.filepath = path

    def run(self):
        self.signOut.emit('程序开始处理', 90)
        manage = ProcessData(self.filepath)

        result = manage.read_variable()

        self.signOut.emit(result[1], 90)
        if result[0]:
            manage.main_func()
            manage.write_data_to_sheet()
            self.signOut.emit('处理完成!', 100)


class QRCodeTransfer(QMainWindow, Surface.Ui_MainWindow):
    """
    Interface the user watch
    """

    def __init__(self):
        QMainWindow.__init__(self)
        Surface.Ui_MainWindow.__init__(self)
        self.setupUi(self)

        self.message = []
        self.slm = QStringListModel()

        self.work_thread = None
        self.step = 0  # 进度条的值
        self.progressBar_Progress.setValue(0)
        self.setWindowIcon(QIcon('./icon.ico'))
        self.Button_SelectExcelFile.clicked.connect(self.select_excel_file)
        self.Button_Start.clicked.connect(self.start_process)

    def select_excel_file(self):
        filename_choose, file_type = QFileDialog.getOpenFileName(self, '打开',
                                                                 os.path.join(os.path.expanduser("~"), 'Desktop'),
                                                                 'Excel文件 (*.xlsm);;All Files (*)')
        self.lineEdit_SelectExcelFile.setText(filename_choose)

    def start_process(self):
        if not self.lineEdit_SelectExcelFile.text():
            QMessageBox.information(self, '提示', '请选择文件或者文件夹!')
        else:
            pass
        self.work_thread = ThreadTransfer(self.lineEdit_SelectExcelFile.text())

        self.work_thread.signOut.connect(self.list_add)
        self.Button_Start.setEnabled(False)
        self.Button_Start.setText('正在处理')
        self.work_thread.start()

    def list_add(self, message, statu):
        self.message.append(message)
        self.slm.setStringList(self.message)
        self.listView_Info.setModel(self.slm)
        self.listView_Info.scrollToBottom()
        self.progressBar_Progress.setValue(statu)
        if statu >= 100:
            self.Button_Start.setEnabled(True)
            self.Button_Start.setText('开始处理')
            QMessageBox.information(self, "提示", "程序处理完成")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # MainWindow = QMainWindow()
    ui = QRCodeTransfer()
    # ui.setupUi(MainWindow)
    ui.show()
    sys.exit(app.exec_())
