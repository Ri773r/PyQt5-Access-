# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt5.QtCore import pyqtSlot, pyqtSignal, QThread
from PyQt5.QtWidgets import QMainWindow, QHeaderView, QAbstractItemView
from PyQt5 import QtWidgets
from Ui_mw import Ui_MainWindow
import pypyodbc
import datetime
import traceback
from setting import Dialog
import os, os.path


class MainWindow(QMainWindow, Ui_MainWindow):
    #向show线程传递列表信号
    #sinOut = pyqtSignal(list)
    def __init__(self, parent=None):
        print('init', int(QThread.currentThreadId()))
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.label_4.hide()
        #数据库链接
        self.driver_path = None#'Driver={Microsoft Access Driver (*.mdb,*.accdb)};DBQ=D:\\江志辉\\DelbyMDB\\权属数据.mdb'
        pypyodbc.lowercase = False
        self.connection = None#pypyodbc.win_connect_mdb(self.driver_path)
        self.curser = None#self.connection.cursor()
        #self.showData()
        self.st = ShowThread()
        self.st.showOut.connect(self.showTable)
        #item（表）表更信号 20180803.0:43改信号连接位置，避免重复绑定导致多次查询
        self.treeWidget.currentItemChanged.connect(self.ItemChanged)
        #绑定信号至槽
        #self.sinOut.connect(self.st.table_setting)
        #运算符字典
        self.ysf = {'小于':'<', '等于':'=', '大于':'>', '不小于':'>=', '不等于':'<>', '不大于':'<=', '文本等于':'str=','文本不等于':'str<>','包含':'like', '不包含':'not like'}
        #表格内容        
        self.table_content = None
        #表格行数
        self.table_row =None
        #表格列数
        self.table_column = None
        #可多选
        self.tableWidget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        #表格不可编辑
        self.tableWidget.setEditTriggers(QAbstractItemView.DoubleClicked)
        #行高
        self.tableWidget.verticalHeader().setDefaultSectionSize(25)
        #自适应表格列宽行高
        #加载很慢
        #self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        #self.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        #combobox item 变换信号
        self.comboBox_3.currentTextChanged.connect(self.showComboboxItem)
        
    def showComboboxItem(self, item_text):
        if item_text == '统赋':
            self.label_5.setMaximumSize(16777215, 16777215)
            self.label_9.setMaximumSize(16777215, 16777215)
            self.lineEdit_3.setMaximumSize(16777215, 16777215)
            self.comboBox_4.setMinimumSize(80, 16777215)
            self.comboBox_4.setMaximumSize(80, 16777215)
        else:
            self.label_5.setMaximumSize(0, 16777215)
            self.label_9.setMaximumSize(0, 16777215)
            self.lineEdit_3.setMaximumSize(0, 16777215)
            self.comboBox_4.setMinimumSize(0, 16777215)
            self.comboBox_4.setMaximumSize(0, 16777215)
    
    def setDb(self, driver_path):
        try:
            if self.curser:
                self.curser.close()
            if self.connection:
                self.connection.close()
            self.driver_path = 'Driver={Microsoft Access Driver (*.mdb,*.accdb)};DBQ=' + driver_path
            self.connection = pypyodbc.win_connect_mdb(self.driver_path)
            self.curser =self.connection.cursor()
            self.treeWidget.clear()
            self.tableWidget.clear()
            self.showData()
            self.setWindowTitle('DBMDB - ' + driver_path + ' - V1.4')
            self.show()
        except pypyodbc.ProgrammingError:
            ts = '[Microsoft][ODBC Microsoft Access Driver]不能读取记录；在"MSYSOBJECTS"上没有读取数据权限。请按以下步骤执行：\n1、打开数据库 "文件"—>"选项"—>"当前数据库"—>"导航选项"—>"显示系统对象"，√，"确定";\
            \n2、"文件"—>"管理用户和权限"—>"用户与组权限"—>"对象名称"—>"MSysObjects"—>"权限"—>"读取数据"，√，确定，关闭。'
            QtWidgets.QMessageBox.question(self,'警告',ts,QtWidgets.QMessageBox.Ok)
            
    #显示数据
    def showData(self):
        #列出表名
        #sql = 'SELECT NAME FROM MSYSOBJECTS WHERE TYPE=1 AND FLAGS=0'
        sql = 'SELECT NAME FROM MSysNameMap'
        self.curser.execute(sql)
        ls = self.curser.fetchall()
        ls = list(zip(*ls))[0]
        #放入treeWidget
        for l in ls:
            item = QtWidgets.QTreeWidgetItem(self.treeWidget)
            item.setText(0, l)
            self.treeWidget.addTopLevelItem(item)
    
    def ItemChanged(self, item, preitem):
        try:
            self.tableWidget.clear()
            table_name = item.text(0)
            self.plainTextEdit.appendPlainText(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+' 选择表格'+table_name)
            sql = 'SELECT * FROM ' + table_name
            self.plainTextEdit.appendPlainText(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+' 执行—>'+sql)
            assert Exception
            print(sql)
            self.curser.execute(sql)
            table_ls = self.curser.fetchall()
            if len(table_ls)>0:
                self.label_4.show()
                self.table_row = len(table_ls)
                self.table_column = len(table_ls[0])
                #创建行列
                self.tableWidget.setRowCount(self.table_row)
                self.tableWidget.setColumnCount(self.table_column)
                
                sql = 'SELECT TOP 1 * FROM ' + table_name
                self.curser.execute(sql)
                table_title = list(zip(*self.curser.description))[0]
                #设置标题
                self.tableWidget.setHorizontalHeaderLabels(list(table_title))
                #设置下拉列表
                self.comboBox.clear()
                self.comboBox_2.clear()
                self.comboBox_3.clear()
                
                #table_title_ls = ['']
                #table_title_ls.extend(list(table_title))
                self.comboBox.addItems(list(table_title))
                self.comboBox_2.addItems(['小于', '等于', '大于', '不小于', '不等于', '不大于', '文本等于', '文本不等于','包含', '不包含'])
                self.comboBox_3.addItems(['选择','删除', '统赋'])
                self.comboBox_4.addItems(list(table_title))
            
                #执行sql
                sql = 'SELECT * FROM ' + table_name
                self.curser.execute(sql)
                self.table_content = self.curser.fetchall()
                #print(self.table_content)
                #添加内容
                #self.sinOut.emit([table_content, table_row, table_column])
                self.label_3.setText('加载中...')
                self.st.start()
                #阻塞当前线程，等待子线程完成，不阻塞（wait）主线程将继续运行，后面会直接崩溃
                self.st.wait()
            else:
                self.tableWidget.clear()
        except Exception as e:
            print(e)
            self.plainTextEdit.appendPlainText(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+' 警告：'+str(e))
        
    @pyqtSlot()
    def on_actionlink_triggered(self):
        da.show()
                
    def closeEvent(self, event):
        self.curser.close()
        self.connection.close()
        event.accept()
    
    @pyqtSlot()
    def on_pushButton_clicked(self):
        try:
            #当前选择的字段
            combobox1 = self.comboBox.currentText()
            #当前选择的筛选项
            combobox2 = self.comboBox_2.currentText()
            #当前选择的操作
            combobox3 = self.comboBox_3.currentText()
            #操作的值
            lineEdit_value = self.lineEdit_2.text()
            
            #当前表格
            current_table = self.treeWidget.currentItem().text(0)
            if combobox3 == '选择':
                combobox3 = 'SELECT'
            elif combobox3 == '删除':
                combobox3 = 'DELETE'
            elif combobox3 == '统赋':
                combobox3 = 'UPDATE'
            else:
                pass
            
            #combobox3 + ' ' + current_table + ' SET ' + combobox1 + '=' + lineEdit_value +
            combobox2 = self.ysf[combobox2]
            if combobox2 == 'like' or combobox2 == 'not like':
                sql = ' WHERE ' + combobox1 + ' ' + combobox2  + " '%" + lineEdit_value + "%'"
            elif combobox2=='>' or combobox2=='=' or combobox2=='<' or combobox2=='>=' or combobox2=='<>' or combobox2=='<=':
                sql = ' WHERE ' + combobox1 + ' ' + combobox2  + ' ' + lineEdit_value
            else:
                sql = ' WHERE ' + combobox1 + ' ' + combobox2.replace('str', '')  + " '" + lineEdit_value + "'"
                
            if combobox3 == 'SELECT' or combobox3 == 'DELETE':
                sql = combobox3 + ' * ' + 'FROM ' + current_table + sql
            elif combobox3 == 'UPDATE':
                change_value = self.lineEdit_3.text()
                combobox4 = self.comboBox_4.currentText()
                sql = combobox3 + ' ' + current_table + ' SET ' + combobox4 + '=\'' + change_value + '\'' + sql
            else:
                pass
            
            
            print(sql)
            self.plainTextEdit.appendPlainText(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+' 执行—>'+sql)
            self.curser.execute(sql)
            if combobox3 == 'SELECT':
                self.table_content = self.curser.fetchall()
                if len(self.table_content) > 0:
                    self.table_row = len(self.table_content)
                    self.table_column = len(self.table_content[0])
                    #发射信号
                    #self.sinOut.emit([table_content, table_row, table_column])
                    self.tableWidget.clearContents()
                    self.label_4.show()
                    #开启线程
                    self.st.start()
                    #阻塞当前线程，等待子线程完成，不阻塞（wait）主线程将继续运行，后面会直接崩溃
                    self.st.wait()
                else:
                    self.tableWidget.clearContents()
                    self.lineEdit.setText(str(0))
            elif combobox3 == 'DELETE':
                QtWidgets.QMessageBox.question(self, '选择','是否删除？', QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
                #删除插入没有返回记录集，所以不能fetchall，而需要提交(commit)，否则无法应用到数据库
                self.curser.commit() 
                #删除成功，刷新页面
                self.curser.execute('SELECT * FROM ' + current_table)
                self.table_content = self.curser.fetchall()
                if len(self.table_content) > 0:
                    self.table_row = len(self.table_content)
                    self.table_column = len(self.table_content[0])
                    #发射信号
                    #self.sinOut.emit([table_content, table_row, table_column])
                    self.tableWidget.clearContents()
                    self.label_4.show()
                    #开启线程
                    self.st.start()
                    #阻塞当前线程，等待子线程完成，不阻塞（wait）主线程将继续运行，后面会直接崩溃
                    self.st.wait()
                else:
                    self.tableWidget.clearContents()
                    self.lineEdit.setText(str(0))
            elif combobox3 == 'UPDATE':
                self.curser.commit() 
                self.curser.execute('SELECT * FROM ' + current_table)
                self.table_content = self.curser.fetchall()
                if len(self.table_content) > 0:
                    self.table_row = len(self.table_content)
                    self.table_column = len(self.table_content[0])
                    #发射信号
                    #self.sinOut.emit([table_content, table_row, table_column])
                    self.tableWidget.clearContents()
                    self.label_4.show()
                    #开启线程
                    self.st.start()
                    #阻塞当前线程，等待子线程完成，不阻塞（wait）主线程将继续运行，后面会直接崩溃
                    self.st.wait()
                else:
                    self.tableWidget.clearContents()
                    self.lineEdit.setText(str(0))
        except Exception as e:
            error_msg = traceback.format_exc()
            self.plainTextEdit.appendPlainText(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+' 错误：'+error_msg)
            QtWidgets.QMessageBox.warning(self, '警告',error_msg, QtWidgets.QMessageBox.Yes)
            

    @pyqtSlot()
    def on_pushButton_2_clicked(self):
        #当前表格
        try:
            current_table = self.treeWidget.currentItem().text(0)
            sql = 'SELECT * FROM ' + current_table
            self.plainTextEdit.appendPlainText(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+' 执行—>'+sql)
            self.curser.execute(sql)
            self.table_content = self.curser.fetchall()
            if len(self.table_content) > 0:
                self.table_row = len(self.table_content)
                self.table_column = len(self.table_content[0])
                #发射信号
                #self.sinOut.emit([table_content, table_row, table_column])
                self.tableWidget.clearContents()
                self.label_4.show()
                #开启线程
                self.st.start()
                #阻塞当前线程，等待子线程完成，不阻塞（wait）主线程将继续运行，后面会直接崩溃
                self.st.wait()
            else:
                self.tableWidget.clearContents()
                self.lineEdit.setText(str(0))
        except Exception as e:
            #error_msg = traceback.format_exc()
            self.plainTextEdit.appendPlainText(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+' 错误：当前没有加载表格')
            QtWidgets.QMessageBox.warning(self, '警告','当前没有加载表格', QtWidgets.QMessageBox.Ok)
            
    #2018.08.02.23:59 尝试把展示函数写到UI类，通过子线程信号激活，用主类更新UI 
    #事实证明，该方法可行，解决信号与槽不同线程之间通讯产生的警告（子线程调用主线程的实例）与界面假死
    #2018.08.03 发现此展示数据函数好像还是在主线程中执行
    def showTable(self):
        print('showTable', int(QThread.currentThreadId()))
        self.progressBar.setMaximum(self.table_row)
        self.plainTextEdit.appendPlainText(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+' 加载中...')
        for row in range(self.table_row):
            for column in range(self.table_column):
                content = self.table_content[row][column]
                if isinstance(self.table_content[row][column],datetime.datetime):
                    content = self.table_content[row][column].strftime('%Y/%m/%d')
                if content is not None:
                    content = str(content)
                table_item = QtWidgets.QTableWidgetItem(content)
                self.tableWidget.setItem(row, column, table_item)
            self.progressBar.setValue(row+1)
            self.label_3.setText(str(row+1)+'/'+str(self.table_row))
        #self.sleep(0.5)
        #self.quit()
        self.lineEdit.setText(str(self.table_row))
        self.plainTextEdit.appendPlainText(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+' 加载完成')
        self.label_3.setText('加载完成')
        self.label_4.hide()
    
    @pyqtSlot()
    def on_pushButton_3_clicked(self):
        if self.pushButton_3.text() == '关闭监视窗口':
            self.plainTextEdit.setMaximumSize(0, 0)
            self.pushButton_3.setText('打开监视窗口')
        elif self.pushButton_3.text() == '打开监视窗口':
            self.plainTextEdit.setMaximumSize(16777215, 16777215)
            self.pushButton_3.setText('关闭监视窗口')
        
        
#展示数据线程
class ShowThread(QThread):
    #加载信号
    #子线程run方法内部发射信号激活UI类执行方法更新UI界面
    showOut = pyqtSignal()
    def __init__(self):
        super(ShowThread, self).__init__()
        #self.table_content = None
        #self.table_row = None
        #self.table_column = None
        
    #def table_setting(self, view_ls):
    #    self.table_content = view_ls[0]
    #   self.table_row = view_ls[1]
    #   self.table_column = view_ls[2]
        
    def run(self):
        print('run', int(QThread.currentThreadId()))
        self.showOut.emit()
        print('OK')

    def __del__(self):
        print('del')
        self.wait()
        
    
if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    mw = MainWindow()
    da = Dialog()
    da.sinOut.connect(mw.setDb)
    da.show()
    sys.exit(app.exec_())
