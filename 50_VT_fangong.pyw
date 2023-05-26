import sys
import time
import datetime
import pymssql
import pymysql
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.uic import loadUiType
#import openpyxl
from PyQt5.uic.properties import QtCore, QtGui
from PyQt5 import QtCore, QtGui, QtWidgets
#from dbutil_20 import get_conn_20, close_conn_20
#from dbutil_50 import get_conn_50, close_conn_50

ui, _ = loadUiType('50_VTfan.ui')

class MainApp(QMainWindow, ui):
   
    # 定义构造
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.handle_buttons()

    def handle_buttons(self):
        self.pihao.returnPressed.connect(self.data)
        self.Scan.returnPressed.connect(self.Barcode_Scan)
        #self.hege_btn.clicked.connect(self.tijiao_hege)
        #self.buhege_btn.clicked.connect(self.tijiao_buhege)
        #self.daiding_btn.clicked.connect(self.tijiao_daiding)
        self.daiding_btn_R.clicked.connect(self.tijiao_pao)
        self.table_data.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_data.customContextMenuRequested.connect(self.youjian)

    def data(self):
        pihao = self.pihao.text()
        
        conn50 = pymysql.connect(host='localhost',database='50',user='root',password='123456')
        cur = conn50.cursor()
      
        try:
            sql = "select Opername,pihao,IndexNo,Barcode_1,Barcode_2,InjuryDes,Result,jianyanyuan,Time,Note from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' order by IndexNo"
            
            cur.execute(sql)
            table_data = cur.fetchall()
  
            row = len(table_data)
            vol = len(table_data[0])

            self.table_data.setRowCount(row)
            self.table_data.setColumnCount(vol)
            #QTableWidgetItem(str(table_data[8][6])).setBackground(QBrush(QColor("red")))
            for i in range(row):
                for j in range(vol):
                    data = QTableWidgetItem(str(table_data[i][j]))
                    data_1 = QTableWidgetItem(str(table_data[i][6])).text()

                    if  data_1 == '合格':
                        data.setBackground(QBrush(QColor("limegreen")))
                    if  data_1 == '不合格':
                        data.setBackground(QBrush(QColor("tomato")))
                    if  data_1 == '待定':
                        data.setBackground(QBrush(QColor("gold")))
                    
                    self.table_data.setItem(i,j,data)
                    data.setTextAlignment(Qt.AlignCenter)
                
            #self.table_data.resizeColumnsToContents()
            self.table_data.resizeRowsToContents()
            self.table_data.setAlternatingRowColors(True)
            self.table_data.scrollToBottom()
            conn50.close()
            #self.yueshu()
            self.tongji()
        except:
            QMessageBox.information(None, '错误', '无数据或预览出错,请检查服务器通讯！')
            conn50.close()

    def tongji(self):
        pihao = self.pihao.text()
        Time = time.strftime("%Y-%m-%d",time.localtime())
        conn50 = pymysql.connect(host='localhost',database='50',user='root',password='123456')
        cur = conn50.cursor()

        sql0 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测'"
        cur.execute(sql0)
        zs = cur.fetchall()
        for z in zs:
            z = list(z)
            z = z[0]
        self.LL.setText("来料："+str(z)+"支")

        sql1 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' and Result = '合格'"
        cur.execute(sql1)
        hs = cur.fetchall()
        for h in hs:
            h = list(h)
            h = h[0]
        self.HG.setText("合格："+str(h)+"支")

        sql2 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' and Result = '不合格'"
        cur.execute(sql2)
        bs = cur.fetchall()
        for b in bs:
            b = list(b)
            b = b[0]
        self.BHG.setText("不合格："+str(b)+"支")

        sql21 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' and Result = '不合格' and InjuryDes like '%外%'"
        cur.execute(sql21)
        bas = cur.fetchall()
        for ba in bas:
            ba = list(ba)
            ba = ba[0]
        self.BHG_W.setText("外表面："+str(ba)+"支")

        sql22 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' and Result = '不合格' and InjuryDes like '%内%'"
        cur.execute(sql22)
        bbs = cur.fetchall()
        for bb in bbs:
            bb = list(bb)
            bb = bb[0]
        self.BHG_N.setText("内表面："+str(bb)+"支")


        sql000 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' and Time like '%"+str(Time)+"%'"
        cur.execute(sql000)
        zs = cur.fetchall()
        for z in zs:
            z = list(z)
            z = z[0]
        self.LL_2.setText("来料："+str(z)+"支")


        sql111 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' and Result = '合格' and Time like '%"+str(Time)+"%'"
        cur.execute(sql111)
        hs = cur.fetchall()
        for h in hs:
            h = list(h)
            h = h[0]
        self.HG_2.setText("合格："+str(h)+"支")

        sql222 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' and Result = '不合格' and Time like '%"+str(Time)+"%'"
        cur.execute(sql222)
        bs = cur.fetchall()
        for b in bs:
            b = list(b)
            b = b[0]
        self.BHG_2.setText("不合格："+str(b)+"支")

        sql2111 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' and Result = '不合格' and InjuryDes like '%外%' and Time like '%"+str(Time)+"%'"
        cur.execute(sql2111)
        bas = cur.fetchall()
        for ba in bas:
            ba = list(ba)
            ba = ba[0]
        self.BHG_W_2.setText("外表面："+str(ba)+"支")

        sql2222 = "select COUNT(*) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' and Result = '不合格' and InjuryDes like '%内%' and Time like '%"+str(Time)+"%'"
        cur.execute(sql2222)
        bbs = cur.fetchall()
        for bb in bbs:
            bb = list(bb)
            bb = bb[0]
        self.BHG_N_2.setText("内表面："+str(bb)+"支")

        conn50.close()

    def Barcode_Scan(self):
        global Barcode_1,Barcode_2
        pihao = self.pihao.text()
        
        ji = ['1', '3', '5', '7', '9']
        Barcode = self.Scan.text()
        conn20 = pymssql.connect(host="10.247.6.1",user="sa",password="xbg123!@#",database="XBG",charset="utf8")
        cur = conn20.cursor()

        sql = "select count(*) from Pipe where PipeCode = '"+str(Barcode)+"'"
        cur.execute(sql)
        nums = cur.fetchall()
        for num in nums:
            num = list(num)
            num = num[0]
        conn20.close()
        
        if num != 0:
            conn50 = pymysql.connect(host='localhost',database='50',user='root',password='123456')
            cur = conn50.cursor()

            sql = "select count(*) from Report where (Result='待定' and Barcode_1='"+str(Barcode)+"' and Opername='初始宏观检测') or (Result='待定' and Barcode_2='"+str(Barcode)+"' and Opername='初始宏观检测')"
            cur.execute(sql)
            dums = cur.fetchall()
            for dum in dums:
                dum = list(dum)
                dum = dum[0]
            conn50.close()

            if dum != 0:
                conn50 = pymysql.connect(host='localhost',database='50',user='root',password='123456')
                cur = conn50.cursor()
                sql = "select count(*) from Report where (Opername='返工宏观检测' and Barcode_1='"+str(Barcode)+"') or (Opername='返工宏观检测' and Barcode_2='"+str(Barcode)+"')"
                cur.execute(sql)
                lums = cur.fetchall()
                for lum in lums:
                    lum = list(lum)
                    lum = lum[0]
                conn50.close()

                if lum == 0:
                    if list(Barcode)[9] in ji:
                        
                        Barcode_1 =  str(Barcode)
                        Barcode_2 = str(int(Barcode)+10)
                    else:
                        Barcode_2 =  str(Barcode)
                        Barcode_1 = str(int(Barcode)-10)

                    Barcode_all = str(Barcode_1)+' '+str(Barcode_2)

                    self.show_Barcode.setText(Barcode_all)

                    self.tijiao()
                    self.Scan.clear()
                else:
                    QMessageBox.information(None, '错误', '管号重复或已完成宏观返工检测！')
                    self.Scan.clear()
            else:
                QMessageBox.information(None, '错误', '非返工管号！')
                self.Scan.clear()
        else:
            QMessageBox.information(None, '错误', '管号非法或不存在！')
            self.Scan.clear()

    def tijiao(self):
        
        pihao = self.pihao.text()
        jianyanyuan = self.jianyanyuan.text()
        #banci = self.banci.currentText()
        biaozhun = self.biaozhun.text()
        shangqing = self.shangqing.text()
        
        #local_time = time.strftime("%Y-%m-%d",time.localtime())
        #Time = local_time+" "+str(banci)

        Time = time.strftime("%Y-%m-%d %H:%M:%S",time.localtime())
        try:
            conn50 = pymysql.connect(host='localhost',database='50',user='root',password='123456')
            cur = conn50.cursor()
        
            cur.execute("select max(IndexNo) from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测'")
            maxs = cur.fetchall()

            for max_ in maxs:
                max_ = list(max_)
                max_ = max_[0]

            if max_ is not None:
                IndexNo = str(int(max_)+1).zfill(3)
            else:
                IndexNo = "001"

            if self.hege_btn.isChecked() is True:
                sql = "insert into Report (Opername,pihao,IndexNo,Barcode_1,Barcode_2,InjuryDes,Result,jianyanyuan,Time,Note) values ('返工宏观检测','"+str(pihao)+"','"+str(IndexNo)+"','"+str(Barcode_1)+"','"+str(Barcode_2)+"','无超标缺陷 尺寸合格','合格','"+str(jianyanyuan)+"','"+str(Time)+"','"+str(biaozhun)+"')"

             
            if self.buhege_btn.isChecked() is True:
        
                sql = "insert into Report (Opername,pihao,IndexNo,Barcode_1,Barcode_2,InjuryDes,Result,jianyanyuan,Time,Note) values ('返工宏观检测','"+str(pihao)+"','"+str(IndexNo)+"','"+str(Barcode_1)+"','"+str(Barcode_2)+"','"+str(shangqing)+"','不合格','"+str(jianyanyuan)+"','"+str(Time)+"','"+str(biaozhun)+"')"
                sql6 = "insert into Report (Opername,pihao,IndexNo,Barcode_1,Barcode_2,InjuryDes,Result,jianyanyuan,Time,Note) values ('返工宏观检测','"+str(pihao)+"','"+str(IndexNo)+"','"+str(Barcode_1)+"','"+str(Barcode_2)+"','无超标缺陷 尺寸合格','合格','"+str(jianyanyuan)+"','"+str(Time)+"','"+str(biaozhun)+"')"


            cur.execute(sql)
            conn50.commit()
            self.data()
            self.show_Barcode.clear()  
            self.shangqing.clear()       
            self.biaozhun.clear()        
            conn50.close()
         
        except:
            QMessageBox.information(None, '错误', '管号提交有误！')
            self.show_Barcode.clear()
            self.biaozhun.clear()
            self.shangqing.clear()
            conn50.close()


    def pao(self):
        global Pipe_all
        Pipe_all = []
        pihao = self.pihao.text()

        conn50 = pymysql.connect(host='localhost',database='50',user='root',password='123456')
        cur = conn50.cursor()

        sql = "select Barcode_1,Barcode_2 from Report where pihao = '" + str(pihao) + "' and Opername like '%宏观%' and Result = '合格'"

        cur.execute(sql)
        Pipe = cur.fetchall()
       
        for i in list(Pipe):
            Pipe_all.append(i[0])
            Pipe_all.append(i[1])
        conn50.close()

    def tijiao_pao(self):
        self.pao()
    
        pihao = self.pihao.text()
        conn20 = pymssql.connect(host="10.247.6.1",user="sa",password="xbg123!@#",database="XBG",charset="utf8")
        cur = conn20.cursor()

        #重复提交？
        
        try:
            for i in Pipe_all:
             
                #sql1 = "insert into Rework (BN,PipeCode,reason_5,reason_9,rework_2,rework_3,Operator_AID,CurPFID,RWState,SN,RWDID,FileIndex) values ('"+str(pihao)+"','"+str(i)+"','1','0','1','1','28','2','0','XXX','XXX','')"
                sql = "insert into Pipe (PipeCode ,Operator_AID,SteelNo,BN,CN,Current_WorkKind,PipeState) values('" + str(i) + "','34','4','" + str(pihao) + "','0','8','2')"
                
                cur.execute(sql)
                conn20.commit()
                
            QMessageBox.information(None, '成功', '管号成功导入包装工序！')
            conn20.close()
     
        except:
            QMessageBox.information(None, '错误', '管号导入错误！')
            conn20.close()
        

    def youjian(self,pos):

        for i in self.table_data.selectionModel().selection().indexes():
            row_num = i.row()

        menu = QMenu()
        item1 = menu.addAction("删除本条记录")
        item2 = menu.addAction("改为合格")
        item3 = menu.addAction("改为不合格")
    
        action = menu.exec_(self.table_data.mapToGlobal(pos))

        if action == item1:
            pihao = self.pihao.text()
            conn50 = pymysql.connect(host='localhost',database='50',user='root',password='123456')
            cur = conn50.cursor()
            sql = "select Opername,pihao,IndexNo,Barcode_1,Barcode_2,InjuryDes,Result,jianyanyuan,Time,Note from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' order by IndexNo"
            cur.execute(sql)
            table_data = cur.fetchall()
            dele = table_data[row_num][2]
            dele = str(int(dele)).zfill(3)
            sql = "delete from Report where pihao = '"+str(pihao)+"' and IndexNo = '"+str(dele)+"' and Opername= '返工宏观检测' "
            cur.execute(sql)
            conn50.commit()
            conn50.close()
            self.data()

        if action == item2:
            pihao = self.pihao.text()  
            conn50 = pymysql.connect(host='localhost',database='50',user='root',password='123456')
            cur = conn50.cursor()
            sql = "select Opername,pihao,IndexNo,Barcode_1,Barcode_2,InjuryDes,Result,jianyanyuan,Time,Note from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' order by IndexNo"         
            cur.execute(sql)
            table_data = cur.fetchall()
            upd = table_data[row_num][2]         
            upd = str(int(upd)).zfill(3)              
            sql = "update Report set  Result = '合格',InjuryDes='无超标缺陷' where pihao = '"+str(pihao)+"' and IndexNo = '"+str(upd)+"' and Opername= '返工宏观检测'"
            cur.execute(sql)
            conn50.commit()
            conn50.close()
            self.show_Barcode.clear()
            self.shangqing.clear()
            self.biaozhun.clear()
            self.data()

        if action == item3:
            pihao = self.pihao.text()
            shangqing = self.shangqing.text()
            conn50 = pymysql.connect(host='localhost',database='50',user='root',password='123456')
            cur = conn50.cursor()
            sql = "select Opername,pihao,IndexNo,Barcode_1,Barcode_2,InjuryDes,Result,jianyanyuan,Time,Note from Report where pihao = '"+str(pihao)+"' and Opername= '返工宏观检测' order by IndexNo"        
            cur.execute(sql)
            table_data = cur.fetchall()
            upd = table_data[row_num][2]            
            upd = str(int(upd)).zfill(3)           
            sql = "update Report set  Result = '不合格',InjuryDes='"+str(shangqing)+"' where pihao = '"+str(pihao)+"' and IndexNo = '"+str(upd)+"' and Opername= '返工宏观检测'"
            cur.execute(sql)
            conn50.commit()
            conn50.close()
            self.show_Barcode.clear()
            self.shangqing.clear()
            self.biaozhun.clear()
            self.data()

def main():
    app = QApplication(sys.argv)
    splash = QSplashScreen(QPixmap("4.png"))
    splash.setFont(QFont('Microsoft YaHei UI', 15))
    splash.show()
    splash.showMessage("启动中...",QtCore.Qt.AlignHCenter |QtCore.Qt.AlignBottom, QtCore.Qt.white)
    time.sleep(0.2)
    ###splash.showMessage("正在连接数据库...",QtCore.Qt.AlignHCenter |QtCore.Qt.AlignBottom, QtCore.Qt.white)
    ###time.sleep(0.2)
    ###conn50 = get_conn_50()
    ###if (conn50 != 0):
     ###   conn50.close()
      ###  splash.showMessage("数据库连接成功...",QtCore.Qt.AlignHCenter |QtCore.Qt.AlignBottom, QtCore.Qt.white)
   ### else:
      ###  splash.showMessage("数据库连接失败...",QtCore.Qt.AlignHCenter |QtCore.Qt.AlignBottom, QtCore.Qt.white)
    # time.sleep(0.5)
    #app = QApplication([])
    window = MainApp()
    window.show()
    splash.finish(window)
    app.exec_()

if __name__ == '__main__':
    main()

