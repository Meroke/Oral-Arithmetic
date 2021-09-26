# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.py'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QWidget, QDesktopWidget
import init2
from functools import partial
import ExcelRead
import os
import traceback

# 返回各类题数，此处为提示‘少于 xx’ ，返回数量比实际大1
def get_AllNums():
    sheet1, sheet2, sheet3 = ExcelRead.get_Allsheets()
    row1, col1, _col1_endlen = ExcelRead.get_line(sheet1)
    row2, col2, _col2_endlen = ExcelRead.get_line(sheet2)
    row3, col3, _col3_endlen = ExcelRead.get_line(sheet3)
    return (row1 + 1) * col1  + _col1_endlen + 2, (row2 + 1) * col2 + _col2_endlen + 2, (row3 + 1) * col3 + _col3_endlen + 2


class MyWindow(init2.Ui_MainWindow):
    def __init__(self, Dialog):
        super().setupUi(Dialog)
        self.YesButton.clicked.connect(partial(self.click_success, self))
        self.NoButton.clicked.connect(self.btnExit)
        self.actionabout.triggered.connect(self.explainMessage)
        self.mul_line.setText(str('40'))
        self.div_line.setText(str('30'))
        self.mix_line.setText(str('30'))
        num1, num2, num3 = get_AllNums()
        self.label_5.setText("少于" + str(num1))
        self.label_6.setText("少于" + str(num2))
        self.label_7.setText("少于" + str(num3))

    def explainMessage(self):
        new = QWidget()
        msg_box = QMessageBox.about(new, '相关信息', '将”三下口算.xlsx“文件放入该程序的同一目录下，即可正常使用。输入题数，请勿超出提示值！')
        # msg_box.exec_()

    def messageWaring(self):
        msg_box = QMessageBox(QMessageBox.Warning, "警告", '请勿输入超出范围')
        msg_box.exec_()

    def messageWaring2(self):
        msg_box = QMessageBox(QMessageBox.Warning, "警告", '请重新启动程序')
        msg_box.exec_()

    def messageWaring3(self):
        msg_box = QMessageBox(QMessageBox.Warning, "警告", '当前目录下缺少“三下口算.xlsx"文件')
        msg_box.exec_()


    def messageWaring4(self):
        msg_box = QMessageBox(QMessageBox.Warning, "警告", '新生成文件可能与已打开文件{}重名，请确认关闭，再重启程序'.format(ExcelRead.create_file_name))
        msg_box.exec_()

    def messageInformation(self):
        msg_box = QMessageBox(QMessageBox.Information, "提示", "测试题成功生成 {}".format(ExcelRead.create_file_name))
        msg_box.exec_()

    def click_success(self, ui):
        if os.path.exists('./三下口算.xlsx'):
            ui.YesButton.setDisabled(True)
            mul_num = ui.mul_line.text()  # 乘法题的数量
            div_num = ui.div_line.text()  # 除法题的数量
            mix_num = ui.mix_line.text()  # 混合题的数量
            if mul_num and div_num and mix_num:  # 全不为0
                try:
                    num1, num2, num3 = get_AllNums()
                    ui.lineEdit.setText(str(int(mul_num) + int(div_num) + int(mix_num)))
                    if 0 <= int(mul_num) < num1 and 0 <= int(div_num) < num2 and 0 <= int(mix_num) < num3:
                        ExcelRead.create_new_file(int(mul_num), int(div_num), int(mix_num))
                        if ExcelRead.file_check:
                            self.messageInformation()
                            ui.YesButton.setDisabled(False)
                        else:
                            self.messageWaring4()
                            ui.YesButton.setDisabled(False)

                    else:
                        self.messageWaring()
                        ui.YesButton.setDisabled(False)
                except Exception as e:
                    traceback.print_exc()
        else:
            self.messageWaring3()
        # else:
        #     self.messageWaring2()

    def btnExit(self):
        sys.exit(app.exec_())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = MyWindow(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
