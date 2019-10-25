__author__ = 'Mahmoud Ahmed'
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4 import QtSql
from PyQt4.uic import loadUiType
import sys
from os import path
#import pypyodbc
#from xlwt import Workbook , Formula , easyxf
#import xlrd
from os import *


FORM_CLASS,_=loadUiType(path.join(path.dirname(__file__),"main.ui"))
FORM_CLASS2,_=loadUiType(path.join(path.dirname(__file__),"login.ui"))
user_profile = []

class login(QWidget , FORM_CLASS2):
    def __init__(self, parent=None):
        super(login,self).__init__(parent)
        QWidget.__init__(self)
        self.setupUi(self)
        self.setFixedSize(600,270)
        self.pushButton.clicked.connect(self.handel_login)
        self.window2 = None

    def handel_login(self):
        self.connection = pypyodbc.connect('DRIVER={SQL Server};SERVER=.\SQLEXPRESS;DATABASE=Cars_DB;')
        self.cursor = self.connection.cursor()
        sql = "SELECT User_Name , Password FROM tbl_Users"
        for row in self.cursor.execute(sql):

            if row[0] == self.lineEdit.text() and row[1] == self.lineEdit_2.text() :
                user_profile.append(self.lineEdit.text())
                user_profile.append(self.lineEdit_2.text())
                print(user_profile)
                self.window2 = main(self)
                self.close()
                self.window2.show()
            else:
                print('error')
                self.label_3.setText("حدث خطأ.. يرجي التأكد من اسم المستخدم وكلمة المرور الخاصة بك ")

    ##################################################################################
    ##################################################################################

class main(QMainWindow,FORM_CLASS):
    def __init__(self,parent=None):
        super(main,self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.initUI()
        self.Handle_Buttons()
        #self.Load_User_Profile()

    def initUI(self):
        self.setWindowTitle('Cars')
        self.setFixedSize(1175,650)
        self.tabWidget.tabBar().setVisible(False)
        self.listWidget.setCurrentRow(0)
        self.radioButton.toggled.connect(self.groupBox_10.setEnabled)
        self.radioButton_5.toggled.connect(self.groupBox_70.setEnabled)
        self.radioButton_4.toggled.connect(self.groupBox_22.setEnabled)
        self.radioButton_8.toggled.connect(self.groupBox_26.setEnabled)
        self.radioButton_10.toggled.connect(self.groupBox_29.setEnabled)
        self.radioButton_11.toggled.connect(self.groupBox_31.setEnabled)
        self.radioButton_13.toggled.connect(self.groupBox_32.setEnabled)
        self.radioButton_16.toggled.connect(self.groupBox_34.setEnabled)

        self.radioButton_4.toggled.connect(self.pushButton_13.setEnabled)
        self.radioButton.toggled.connect(self.pushButton_14.setEnabled)
        self.radioButton_5.toggled.connect(self.pushButton_15.setEnabled)
        self.radioButton_5.toggled.connect(self.pushButton_16.setEnabled)
        self.radioButton_8.toggled.connect(self.pushButton_17.setEnabled)
        self.radioButton_8.toggled.connect(self.pushButton_18.setEnabled)
        self.radioButton_10.toggled.connect(self.btn_save_21.setEnabled)
        self.radioButton_11.toggled.connect(self.btn_save_22.setEnabled)
        self.radioButton_13.toggled.connect(self.btn_save_23.setEnabled)
        self.radioButton_17.toggled.connect(self.btn_save_24.setEnabled)


        #self.connection = pypyodbc.connect('DRIVER={SQL Server};SERVER=.\SQLEXPRESS;DATABASE=Cars_DB;')
        #self.cursor = self.connection.cursor()

        self.dateEdit_5.setDateTime(QDateTime.currentDateTime())
        self.dateEdit.setDateTime(QDateTime.currentDateTime())
        self.dateEdit_2.setDateTime(QDateTime.currentDateTime())
        self.dateEdit_6.setDateTime(QDateTime.currentDateTime())
        self.dateEdit_7.setDateTime(QDateTime.currentDateTime())
        self.tabWidget.setCurrentIndex(0)
        self.tabWidget_2.setCurrentIndex(0)


    def Handle_Buttons(self):
        self.btn_save.clicked.connect(self.Save_Car_Data)
        self.btn_save_2.clicked.connect(self.Save_Fuell_Data)
        self.pushButton_6.clicked.connect(self.save_Re_Maintenance_Data)
        self.btn_save_11.clicked.connect(self.Re_Licenses_Data)
        self.btn_save_12.clicked.connect(self.Save_Un_Re_Licenses_Data)
        self.btn_save_5.clicked.connect(self.Save_Rents_Data)
        self.btn_save_4.clicked.connect(self.save_Revenu_Data)
        self.btn_save_6.clicked.connect(self.Save_Alclarkat_Data)
        self.btn_save_7.clicked.connect(self.Save_EL_WA_Data)
        self.pushButton_3.clicked.connect(self.Car_choice)
        self.btn_save_13.clicked.connect(self.Revenu_choice)
        self.bt_search.clicked.connect(self.search_click)
        self.pushButton.clicked.connect(self.Fuel_choice)
        self.pushButton_4.clicked.connect(self.Maintanance_choice)
        self.pushButton_7.clicked.connect(self.Licence_choice)
        self.btn_save_14.clicked.connect(self.Rents_choice)
        self.btn_save_17.clicked.connect(self.Water_Electricty_choice)
        self.btn_save_3.clicked.connect(self.delete_car)
        self.pushButton_10.clicked.connect(self.UN_R_Licence_choice)
        self.pushButton_9.clicked.connect(self.save_UN_Re_Maintenance_Data)
        self.btn_save_19.clicked.connect(self.add_user)
        self.btn_save_9.clicked.connect(self.add_user_privelige)
        self.btn_save_18.clicked.connect(self.update_user_password)
        self.pushButton_11.clicked.connect(self.Un_R_Maintanance_choice)
        self.btn_save_15.clicked.connect(self.Clarkat_choice)
        self.btn_save_20.clicked.connect(self.update_car_data)


        self.pushButton_13.clicked.connect(self.Retrieve_Car_Data)
        self.pushButton_14.clicked.connect(self.Retrieve_Fuel_Data)
        self.pushButton_15.clicked.connect(self.Retrieve_R_Maintenance_Data)
        self.pushButton_16.clicked.connect(self.Retrieve_Un_R_Maintenance_Data)
        self.pushButton_17.clicked.connect(self.Retrieve_R_Licence_Data)
        self.pushButton_18.clicked.connect(self.Retrieve_Un_R_Licence_Data)
        self.btn_save_21.clicked.connect(self.Retrieve_Revenu_Data)
        self.btn_save_22.clicked.connect(self.Retrieve_Rents_Data)
        self.btn_save_23.clicked.connect(self.Retrieve_Clarkat_Data)
        self.btn_save_24.clicked.connect(self.Retrieve_Weter_Electricty_Data)
        self.pushButton_12.clicked.connect(self.Retrieve_Un_R_Maintenance_Data)

    def Load_User_Profile(self):
        sql = "SELECT * FROM tbl_Users"
        for row in self.cursor.execute(sql):

            if row[0] == user_profile[0] and row[1] == user_profile[1] :
                print(row)
                #self.btn_save.setEnabled(False) ; self.btn_save_3.setEnabled(False) if row[2] == 0 else  print('error')
                if row[2] == 0 :
                    self.btn_save.setEnabled(False)
                    self.btn_save_3.setEnabled(False)
                    self.label_14.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')

                #self.btn_save_11.setEnabled(False) ; self.btn_save_12.setEnabled(False) if row[3] == 0 else  print('error')
                if row[3] == 0 :
                    self.btn_save_11.setEnabled(False)
                    self.btn_save_12.setEnabled(False)
                    self.label_34.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')

                #self.btn_save_5.setEnabled(False) if row[4] == 0 else  print('error')
                if row[4] == 0 :
                    self.btn_save_5.setEnabled(False)
                    self.label_66.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')

                #self.btn_save_2.setEnabled(False) if row[5] == 0 else  print('error')
                if row[5] == 0 :
                    self.btn_save_2.setEnabled(False)
                    self.label_26.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')


                #self.btn_save_4.setEnabled(False) if row[6] == 0 else  print('error')
                if row[6] == 0 :
                    self.btn_save_4.setEnabled(False)
                    self.label_50.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')


                #self.btn_save_6.setEnabled(False) if row[7] == 0 else  print('error')
                if row[7] == 0 :
                     self.btn_save_6.setEnabled(False)
                     self.label_138.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')


                #self.pushButton_6.setEnabled(False) ; self.pushButton_9.setEnabled(False) if row[3] == 0 else  print('error')
                if row[8] == 0 :
                    self.pushButton_6.setEnabled(False)
                    self.pushButton_9.setEnabled(False)
                    self.label_33.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')


                #self.btn_save_7.setEnabled(False) if row[9] == 0 else  print('error')
                if row[9] == 0 :
                    self.btn_save_7.setEnabled(False)
                    self.label_139.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')


                #self.pushButton_3.setEnabled(False) if row[10] == 0 else  print('error')
                if row[10] == 0 :
                    self.pushButton_3.setEnabled(False)
                    self.label_141.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')

                #self.btn_save_8.setEnabled(False) if row[11] == 0 else  print('error')
                if row[11] == 0  :
                    self.btn_save_8.setEnabled(False)
                    self.label_140.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')

                #self.btn_save_19.setEnabled(False) ; self.btn_save_9.setEnabled(False);self.btn_save_18.setEnabled(False) ; self.btn_save_10.setEnabled(False) if row[12] == 0 else  print('error')
                if row[12] == 0 :
                    self.btn_save_19.setEnabled(False)
                    self.btn_save_9.setEnabled(False)
                    self.btn_save_18.setEnabled(False)
                    self.btn_save_10.setEnabled(False)
                    self.label_142.setText('ليس لديك الصلاجية للتعامل مع هذة الواجهة')


    def search_click(self):
        print(self.listWidget.currentItem().text())
        print(self.listWidget.row(self.listWidget.currentItem()))

        if self.listWidget.row(self.listWidget.currentItem()) == 0 :
            print('index 0')
            sql = "SELECT * FROM tbl_Car WHERE Car_Num = ?"
            car_number = int(self.le_search.text())

            for row in self.cursor.execute(sql , [(car_number)]):
                print(str(row))
                self.le_car_number.setText(str(row[1]))
                self.le_owner_company.setText(row[2])
                self.le_branch.setText(row[3])
                self.cm_service_mode.setCurrentIndex(int(row[4]))
                self.le_shaceh_number.setText(row[5])
                self.le_motor_number.setText(row[6])
                self.cm_full_kind.setCurrentIndex(int(row[7]))
                self.le_car_kind.setText(row[8])
                self.le_car_model.setText(str(row[9]))
                self.le_car_load.setText(str(row[10]))
                self.le_car_weight.setText(str(row[11]))
                self.le_shape.setText(row[12])
                self.le_color.setText(row[13])

    ########################################################################
        if self.listWidget.row(self.listWidget.currentItem()) == 1 :
            print('index 1')
            sql = "SELECT * FROM tbl_Fuel WHERE Car_Num = ?"
            car_number = int(self.le_search.text())

            for row in self.cursor.execute(sql , [(car_number)]):
                print(str(row))
                self.le_leters_number.setText(str(row[2]))
                self.le_price.setText(str(row[3]))
                self.le_meter_number.setText(str(row[4]))
                self.lcd_number.display(row[5])

    ###########################################################################
        if self.listWidget.row(self.listWidget.currentItem()) == 2 :
            print('index 2')
            sql = "SELECT * FROM tbl_R_Maintanance WHERE Car_Num = ?"
            car_number = int(self.le_search.text())

            for row in self.cursor.execute(sql , [(car_number)]):
                print(str(row))
                self.lineEdit_45.setText(row[1])
                self.lineEdit_47.setText(row[2])
                self.lineEdit_44.setText(row[3])
                self.lineEdit_46.setText(row[4])
                self.lineEdit_48.setText(row[5])
                self.lineEdit_50.setText(row[6])
                self.lineEdit_43.setText(row[7])
                self.lineEdit_49.setText(row[8])
                self.pt_1.setText(row[9])

    ###########################################################################
        if self.listWidget.row(self.listWidget.currentItem()) == 3 :
            print('index 3')
            sql = "SELECT * FROM tbl_Re_Licences WHERE Car_Num = ?"
            car_number = int(self.le_search.text())

            for row in self.cursor.execute(sql , [(car_number)]):
                print(str(row))
                self.lineEdit_32.setText(row[1])
                self.lineEdit_37.setText(row[2])
                self.lineEdit_40.setText(row[3])
                self.lineEdit_38.setText(row[4])
                self.lineEdit_39.setText(row[5])
                self.lineEdit_41.setText(row[6])
                self.lineEdit_42.setText(row[7])
                self.lineEdit_36.setText(row[8])

    #############################################################################
        if self.listWidget.row(self.listWidget.currentItem()) == 4 :
            print('index 4')
            sql = "SELECT * FROM tbl_Revnue WHERE Car_Num = ?"
            car_number = int(self.le_search.text())

            for row in self.cursor.execute(sql , [(car_number)]):
                print(str(row))
                self.le_revenues_value_6.setText(str(row[2]))
                self.le_revenues_value_5.setText(str(row[3]))
                self.le_revenues_value_7.setText(str(row[4]))

    ############################################################################
        if self.listWidget.row(self.listWidget.currentItem()) == 5 :
            print('index 5')
            sql = "SELECT * FROM tbl_Rents WHERE Car_Num = ?"
            car_number = int(self.le_search.text())

            for row in self.cursor.execute(sql , [(car_number)]):
                print(str(row))
                self.le_rents_value_4.setText(row[2])
                self.le_rents_value_5.setText(row[3])
    ##############################################################################3
        if self.listWidget.row(self.listWidget.currentItem()) == 6 :
            print('index 6')
            sql = "SELECT * FROM tbl_Clarkat WHERE Car_Num = ?"
            car_number = int(self.le_search.text())

            for row in self.cursor.execute(sql , [(car_number)]):
                print(str(row))
                self.le_kind_value_4.setText(row[1])
                self.le_kind_value_5.setText(row[2])
    #############################################################################
        if self.listWidget.row(self.listWidget.currentItem()) == 7 :
            print('index 7')
            sql = "SELECT * FROM tbl_Water_Electricty WHERE Car_Num = ?"
            car_number = int(self.le_search.text())

            for row in self.cursor.execute(sql , [(car_number)]):
                print(str(row))
                self.le_elctric_value.setText(row[1])
                self.le_water_water_value.setText(row[2])
    ###########################################################################
        if self.listWidget.row(self.listWidget.currentItem()) == 8 :
            print('index 8')

    #########################################################################
        if self.listWidget.row(self.listWidget.currentItem()) == 9 :
            print('index 9')


    ############################################################################
        if self.listWidget.row(self.listWidget.currentItem()) == 10 :
            print('index 10')



    def Save_Car_Data(self):
        x = self.le_car_number.text()
        q = self.le_owner_company.text()
        c = self.le_branch.text()

        if self.cm_service_mode.currentIndex() == 0 :
            s = 'بيع'
        elif self.cm_service_mode.currentIndex() == 1 :
            s ='ملاكي'
        elif self.cm_service_mode.currentIndex() == 2 :
            s= 'امداد'
        else:
            s='خدمة مصنع'


        sh = self.le_shaceh_number.text()
        m = self.le_motor_number.text()

        if self.cm_full_kind.currentIndex() == 0 :
            ci = 'سولار'

        else:
            ci = 'بنزين'


        ck = self.le_car_kind.text()
        cm = self.le_car_model.text()
        l = self.le_car_load.text()
        cw = self.le_car_weight.text()
        cs = self.le_shape.text()
        cc = self.le_color.text()
        dd = self.dateEdit_5.text()

        self.cursor.execute("INSERT INTO tbl_Car(Car_Num,Owner_Company,Branch,Service_Mode,Shaceh_Number,Motor_Number,Fuel_Type,Car_Type,Car_Model,Car_Load,Car_Weight,Shape,Color,Add_Date)"
                " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)" , (x,q,c,s,sh,m,ci,ck,cm,l,cw,cs,cc,dd))


        self.cursor.execute("INSERT INTO tbl_Re_Licences(Car_Num) VALUES (?)" ,(self.le_car_number.text(),))
        self.cursor.execute("INSERT INTO tbl_Fuel(Car_Num) VALUES (?)" ,(self.le_car_number.text(),))
        self.cursor.execute("INSERT INTO tbl_R_Maintanance(Car_Num) VALUES (?)" , (self.le_car_number.text(),))
        self.cursor.execute("INSERT INTO tbl_Revnue(Car_Num) VALUES (?)" , (self.le_car_number.text(),))
        self.cursor.execute("INSERT INTO tbl_Clarkat(Car_Num) VALUES (?)" , (self.le_car_number.text(),))
        self.cursor.execute("INSERT INTO tbl_Water_Electricty(Car_Num) VALUES (?)" , (self.le_car_number.text(),))
        self.cursor.execute("INSERT INTO tbl_Rents(Car_Num) VALUES (?)" , (self.le_car_number.text(),))

        self.cursor.commit()
        self.statusbar.showMessage('تم حفظ البيانات بنجاح')
        self.le_car_number.setText('')
        self.le_owner_company.setText('')
        self.le_branch.setText('')
        self.cm_service_mode.setCurrentIndex(0)
        self.le_shaceh_number.setText('')
        self.le_motor_number.setText('')
        self.cm_full_kind.setCurrentIndex(0)
        self.le_car_kind.setText('')
        self.le_car_model.setText('')
        self.le_car_load.setText('')
        self.le_car_weight.setText('')
        self.le_shape.setText('')
        self.le_color.setText('')
        self.statusbar.showMessage('تم الاضافة بنجاح')


    def update_car_data(self):
        q = self.le_owner_company.text()
        c = self.le_branch.text()
        s = int(self.cm_service_mode.currentIndex())
        sh = self.le_shaceh_number.text()
        m = self.le_motor_number.text()
        ci = self.cm_full_kind.currentIndex()
        ck = self.le_car_kind.text()
        cm = self.le_car_model.text()
        l = self.le_car_load.text()
        cw = self.le_car_weight.text()
        cs = self.le_shape.text()
        cc = self.le_color.text()

        self.cursor.execute("UPDATE  tbl_Car SET Owner_Company=?,Branch=?,Service_Mode=?,Shaceh_Number=?,Motor_Number=?,Fuel_Type=?,Car_Type=?,Car_Model=?,Car_Load=?,Car_Weight=?,Shape=?,Color=? WHERE Car_Num= ?"
                , (q,c,int(s),sh,m,ci,ck,cm,l,cw,cs,cc,self.le_search.text()))

        self.cursor.commit()
        self.statusbar.showMessage('تم حفظ البيانات بنجاح')
        self.le_car_number.setText('')
        self.le_owner_company.setText('')
        self.le_branch.setText('')
        self.cm_service_mode.setCurrentIndex(0)
        self.le_shaceh_number.setText('')
        self.le_motor_number.setText('')
        self.cm_full_kind.setCurrentIndex(0)
        self.le_car_kind.setText('')
        self.le_car_model.setText('')
        self.le_car_load.setText('')
        self.le_car_weight.setText('')
        self.le_shape.setText('')
        self.le_color.setText('')
        self.statusbar.showMessage('تم التعديل بنجاح')



    def Save_Fuell_Data(self):
        self.cursor.execute("UPDATE  tbl_Fuel SET Liter_Num = ? ,Meter_Num = ? ,Total = ? ,tips = ? ,price = ? WHERE Car_Num= ? "
                        ,(self.le_leters_number.text(),self.le_meter_number.text(),self.lcd_number.digitCount(),self.le_tips.text(),self.le_price.text(),self.le_search.text()))
        self.cursor.commit()
        self.statusbar.showMessage('تم حفظ البيانات بنجاح')
        self.le_leters_number.setText('')
        self.le_meter_number.setText('')
        self.lcd_number.digitCount(0)
        self.le_tips.setText('')
        self.le_price.setText('')
        print('Done')

    def get_total(self):
        pass


    def save_Re_Maintenance_Data(self):

        self.cursor.execute(" UPDATE tbl_R_Maintanance SET Overhauls_engines = ?,Mechanics = ?,Electricity= ?,Denting_Doku= ? ,Afhh = ? ,Periodic_maintenance= ? ,Cooling = ? , glass = ? , Notes=? WHERE Car_Num= ?"
                    ,(self.lineEdit_45.text(),self.lineEdit_47.text(),self.lineEdit_44.text(),self.lineEdit_46.text(),self.lineEdit_48.text(),self.lineEdit_50.text(),self.lineEdit_43.text(),self.lineEdit_49.text(),self.pt_1.toPlainText(),self.le_search.text()))
        self.cursor.commit()
        self.statusbar.showMessage('تم حفظ البيانات بنجاح')
        self.lineEdit_45.setText('')
        self.lineEdit_47.setText('')
        self.lineEdit_44.setText('')
        self.lineEdit_46.setText('')
        self.lineEdit_48.setText('')
        self.lineEdit_50.setText('')
        self.lineEdit_43.setText('')
        self.lineEdit_49.setText('')
        self.pt_1.toPlainText('')
        self.le_search.setText('')
        print('DOne')


    def save_UN_Re_Maintenance_Data(self):
        print(self.dateEdit.text())
        print(type(self.dateEdit.text()))

        self.cursor.execute("INSERT INTO tbl_Un_Re_Maintainance (Form , [To] , Value) "
                            "VALUES (?,?,?)" ,(self.dateEdit.text(),self.dateEdit_2.text(),self.le_value.text()))
        self.cursor.commit()
        self.statusbar.showMessage('تم حفظ البيانات بنجاح')
        self.dateEdit.setDateTime(QDateTime.currentDateTime())
        self.dateEdit_2.setDateTime(QDateTime.currentDateTime())
        self.le_value.setText('')

    def Re_Licenses_Data(self):
         self.cursor.execute("UPDATE tbl_Re_Licences SET Annual_renewal=?,New_car_license=?,Renew_permit=?,Inquire_irregularities=?,Annual_irregularities=?,Stamp_receipts=?,Fined_drawn_License=?,Instant_fined_irregularities=? WHERE Car_Num= ?)"
                ,(self.lineEdit_32.text(),self.lineEdit_37.text(),self.lineEdit_40.text(),self.lineEdit_38.text(),self.lineEdit_39.text(),self.lineEdit_41.text(),self.lineEdit_42.text(),self.lineEdit_36.text(),self.le_search.text()))

         self.cursor.commit()
         self.statusbar.showMessage('تم حفظ البيانات بنجاح')
         self.lineEdit_32.setText('')
         self.lineEdit_37.setText('')
         self.lineEdit_40.setText('')
         self.lineEdit_38.setText('')
         self.lineEdit_39.setText('')
         self.lineEdit_41.setText('')
         self.lineEdit_42.setText('')
         self.lineEdit_36.setText('')
         self.le_search.setText('')



    def Save_Un_Re_Licenses_Data(self):
         #self.cursor.commit("UPDATE tbl_Un_Re_Licenses SET Data_certificates=? ,Ads_taxes=? ,Health_Traffic_letters=? ,Allowances_transmission ,"
          #                  ,(self.le_licenses_value_11.text(),self.le_licenses_value_10.text(),self.le_licenses_value_9.text(),self.le_licenses_value_8 , self.le_search.text()))

         print(type(self.dateEdit_5.text()))
         self.cursor.execute('INSERT INTO tbl_Un_Re_Licenses(Data_certificates,Ads_taxes,Health_Traffic_letters,Allowances_transmission,Add_Date)'
                             'VALUES(?,?,?,?,?)', (self.le_licenses_value_11.text(),self.le_licenses_value_10.text(),self.le_licenses_value_9.text(),self.le_licenses_value_8.text(),self.dateEdit_5.text()))
         self.cursor.commit()
         self.statusbar.showMessage('تم حفظ البيانات بنجاح')
         self.le_licenses_value_11.setText('')
         self.le_licenses_value_10.setText('')
         self.le_licenses_value_9.setText('')
         self.le_licenses_value_8.setText('')
         self.dateEdit_5.setDateTime(QDateTime.currentDateTime())
         self.le_search.setText('')
         print('Doooonnee')



    def save_Revenu_Data(self):
        self.cursor.execute("UPDATE tbl_Revnue SET Sale_priests=? , Compensation_insurance=? ,Accident_compensation=? WHERE Car_Num= ?"
         ,(self.le_revenues_value_6.text(),self.le_revenues_value_5.text(),self.le_revenues_value_7.text(),self.le_search.text()))
        self.cursor.commit()
        self.statusbar.showMessage('تم حفظ البيانات بنجاح')
        self.le_revenues_value_6.setText('')
        self.le_revenues_value_5.setText('')
        self.le_revenues_value_7.setText('')
        self.le_search.setText('')
        print("Dooooooooooooone")



    def Save_Rents_Data (self):
        self.cursor.execute("UPDATE tbl_Rents SET Monthly_rent=? , Ghafrh_Arabs_Rent=?  WHERE Car_Num= ?"
                            ,(self.le_rents_value_4.text() , self.le_rents_value_5.text(),self.le_search.text()))
        self.cursor.commit()
        self.statusbar.showMessage('تم حفظ البيانات بنجاح')
        self.le_rents_value_4.setText('')
        self.le_rents_value_5.setText('')
        self.le_search.setText('')



    def Save_Alclarkat_Data (self):
        self.cursor.execute("UPDATE tbl_Clarkat SET Rent_Clarkat=? , Maintenance_Clarkat = ? WHERE Car_Num= ?"
                            ,(self.le_kind_value_4.text(),self.le_kind_value_5.text(),self.le_search.text()))
        self.cursor.commit()
        self.le_kind_value_4.setText('')
        self.le_kind_value_5.setText('')
        self.le_search.setText('')
        self.statusbar.showMessage('تم حفظ البيانات بنجاح')



    def Save_EL_WA_Data (self):
        self.cursor.execute("UPDATE tbl_Water_Electricty SET electricity_value= ? , water_value = ? WHERE Car_Num= ? "
                            ,(self.le_elctric_value.text() , self.le_water_water_value.text(),self.le_search.text()))
        self.cursor.commit()
        self.statusbar.showMessage('تم حفظ البيانات بنجاح')
        self.le_elctric_value.setText('')
        self.le_water_water_value.setText('')
        self.le_search.setText('')

    ############################################################################
    ############################################################################
        #  report
        #### cars report

    def one_car_details_report(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet 1')
        sql = "SELECT * From tbl_Car WHERE Car_Num = ?"
        car_number = int(self.le_search.text())
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            print(row[0])
            print(row[4])
            sheet1.col(0).width = 7000
            sheet1.col(0).height_mismatch = True
            sheet1.col(0).height = 20*2566
            sheet1.col(2).width = 7000
            sheet1.col(2).height_mismatch = True
            sheet1.col(2).height = 20*256
            sheet1.write(0,2, "رقم السيارة" ,style1)
            sheet1.write(0,0,row[1] ,style1)

            sheet1.write(1,2, "الشركة المالكة" , style1)
            sheet1.write(1,0,row[2] ,style1)

            sheet1.write(2,2, "الفرع" , style1)
            sheet1.write(2,0,row[3] ,style1)

            sheet1.write(3,2, "فرع الخدمة" , style1) # اظبط الانديكس
            sheet1.write(3,0,row[4] ,style1)

            sheet1.write(4,2, "رقم الشاسية" , style1)
            sheet1.write(4,0,row[5] ,style1)

            sheet1.write(5,2, "رقم الموتور"  , style1)
            sheet1.write(5,0,row[6] ,style1)

            sheet1.write(6,2 , "نوع الوقود" , style1)
            sheet1.write(6,0,row[7] ,style1)

            sheet1.write(7,2 , " نوع السيارة" , style1)
            sheet1.write(7,0,row[8] ,style1)

            sheet1.write(8,2, "موديل السيارة" , style1)
            sheet1.write(8,0,row[9] ,style1)

            sheet1.write(9,2, "حمولة السيارة" , style1)
            sheet1.write(9,0,row[10] ,style1)

            sheet1.write(10,2, "وزن السيارة" , style1)
            sheet1.write(10,0,row[11] ,style1)

            sheet1.write(11,2, "الشكل" , style1)
            sheet1.write(11,0,row[12] ,style1)

            sheet1.write(12,2, "اللون" , style1)
            sheet1.write(12,0,row[13] ,style1)

            sheet1.write(13,2, "تاريخ اضافة السيارة", style1)
            sheet1.write(13,0,row[14],style1)

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/cars.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')



    def all_car_details_report(self):
        self.cursor.execute("SELECT * From tbl_Car")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        print(self.result[1])

        wb = Workbook()
        sheet1 = wb.add_sheet('All Cars')

        sheet1.write(0,1, "رقم السيارة")
        sheet1.write(0,2,"الشركة المالكة")
        sheet1.write(0,3,"الفرع")
        sheet1.write(0,4,"وضع الخدمة")
        sheet1.write(0,5,"رقم الشاسية")
        sheet1.write(0,6,"رقم الموتور")
        sheet1.write(0,7,"نوع وقود")
        sheet1.write(0,8,"نوع سيارة")
        sheet1.write(0,9,"موديل السيارة")
        sheet1.write(0,10,"حمولة السيارة")
        sheet1.write(0,11,"وزن السيارة")
        sheet1.write(0,12,"الشكل ")
        sheet1.write(0,13,"اللون")
        sheet1.write(0,14,"تاريخ اضافة السيارة")

        row_number = 2
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/cars.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')



    def Car_choice(self):
        if self.radioButton_4.isChecked() == True:
            self.one_car_details_report()

        if self.radioButton_3.isChecked() == True:
            self.all_car_details_report()

    #######################################################################
             #### Fuel Report

    def one_car_fuel_report(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet 1')
        sql = "SELECT * From tbl_Fuel WHERE Car_Num = ?"
        car_number = int(self.le_search.text())
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            sheet1.col(0).width = 7000
            sheet1.col(0).height_mismatch = True
            sheet1.col(0).height = 20*2566
            sheet1.col(2).width = 7000
            sheet1.col(2).height_mismatch = True
            sheet1.col(2).height = 20*256
            sheet1.write(0,2, "رقم السيارة" ,style1)
            sheet1.write(0,0,row[1] ,style1)
            sheet1.write(1,2,"عدد اللترات " , style1)
            sheet1.write(1,0,row[2] ,style1)
            sheet1.write(2,2,"سعر اللتر" , style1)
            sheet1.write(2,0 ,row[3] , style1)
            sheet1.write(3,2, "قرأة العداد " , style1)
            sheet1.write(3,0 ,row[4] ,style1)
            sheet1.write(4,2, "السعر الكلي" , style1)
            sheet1.write(4,0 ,row[5] ,style1)
            sheet1.write(5,2, "الاكراميات" , style1)
            sheet1.write(5,0 ,row[6] ,style1)

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/fuel.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')



    def all_car_fuel_report(self):
        self.cursor.execute("SELECT * From tbl_Fuel")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        print(self.result[1])

        wb = Workbook()
        sheet1 = wb.add_sheet('All Cars Fuel Report')

        sheet1.write(0,1, "رقم السيارة")
        sheet1.write(0,2,"عدد اللترات")
        sheet1.write(0,3,"سعر اللتر")
        sheet1.write(0,4,"قراة العداد")
        sheet1.write(0,5,"الاجمالي")
        sheet1.write(0,6,"الاكراميات")


        row_number = 1
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/fuel.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def Fuel_choice(self):
        if self.radioButton.isChecked() == True:
            self.one_car_fuel_report()

        if self.radioButton_2.isChecked() == True:
            self.all_car_fuel_report()

###################################################33
    ## maintanance report

    def one_car_maintanance_report(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('sheet 1')
        sql = "SELECT * From tbl_R_Maintanance WHERE Car_Num = ?"
        car_number = int(self.le_search.text())
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            sheet1.col(0).width = 7000
            sheet1.col(0).height_mismatch = True
            sheet1.col(0).height = 20*2566
            sheet1.col(2).width = 7000
            sheet1.col(2).height_mismatch = True
            sheet1.col(2).height = 20*256
            sheet1.write(0,2, "رقم السيارة" ,style1)
            sheet1.write(0,0,row[0] ,style1)
            sheet1.write(1,2,"عمرات محركات" ,style1)
            sheet1.write(1,0,row[1],style1)
            sheet1.write(2,2,"ميكانيكا",style1)
            sheet1.write(2,0,row[2],style1)
            sheet1.write(3,2,"كهرباء",style1)
            sheet1.write(3,0,row[3],style1)
            sheet1.write(4,2,"سمكر ودوكو",style1)
            sheet1.write(4,0,row[4],style1)
            sheet1.write(5,2,"عفشة",style1)
            sheet1.write(5,0,row[5],style1)
            sheet1.write(6,2,"صيانة دورية بالتوكيلات والورش",style1)
            sheet1.write(6,0,row[6],style1)
            sheet1.write(7,2,"تبريد",style1)
            sheet1.write(7,0,row[7],style1)
            sheet1.write(8,2,"زجاج",style1)
            sheet1.write(8,0,row[8],style1)
            sheet1.write(9,2,"ملاحظات",style1)
            sheet1.write(9,0,row[9],style1)

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/maintainance.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')



    def all_car_maintainance_report(self):
        self.cursor.execute("SELECT * From tbl_R_Maintanance")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        print(self.result[1])

        wb = Workbook()
        sheet1 = wb.add_sheet('All Cars Maintanance Report')

        row_number = 1
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/maintainance.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')



    def Maintanance_choice(self):
        if self.radioButton_5.isChecked() == True:
            self.one_car_maintanance_report()

        if self.radioButton_6.isChecked() == True:
            self.all_car_maintainance_report()

    #####################################################

    def one_car_Un_R_maintanance_report(self):

        wb = Workbook()
        sheet1 = wb.add_sheet('sheet 1')
        sql = "SELECT * From tbl_Un_Re_Maintainance WHERE Form = ? AND [To]= ?"
        max_value = self.dateEdit_6.text()
        minvalue = self.dateEdit_7.text()
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')

        for row in self.cursor.execute(sql,(max_value,minvalue)):

            print(str(row))
            sheet1.col(0).width = 7000
            sheet1.col(0).height_mismatch = True
            sheet1.col(0).height = 20*2566
            sheet1.col(2).width = 7000
            sheet1.col(2).height_mismatch = True
            sheet1.col(2).height = 20*256

            sheet1.write(0,2, " من " ,style1)
            sheet1.write(0,0,row[1] ,style1)
            sheet1.write(1,2,"الي" ,style1)
            sheet1.write(1,0,row[2],style1)
            sheet1.write(2,2,"القيمة",style1)
            sheet1.write(2,0,row[3],style1)


        fileName = QFileDialog.getSaveFileName(self, "Save","E:/maintainance.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')



    def all_car_Un_R_maintainance_report(self):

        self.cursor.execute("SELECT * From tbl_Un_Re_Maintainance")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        print(self.result[1])

        wb = Workbook()
        sheet1 = wb.add_sheet('All Cars Maintanance Report')

        sheet1.write(0,1,"من")
        sheet1.write(0,2,"الي")
        sheet1.write(0,3,"القيمة المالية")

        row_number = 1
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/maintainance.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')



    def Un_R_Maintanance_choice(self):
        if self.radioButton_5.isChecked() == True:
            self.one_car_Un_R_maintanance_report()

        if self.radioButton_6.isChecked() == True:
            self.all_car_Un_R_maintainance_report()

    #####################################################
    ## Relational Licence
    def one_car_licence(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('sheet 1')
        sql = "SELECT * From tbl_Re_Licences WHERE Car_Num = ?"
        car_number = int(self.le_search.text())
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            sheet1.col(0).width = 7000
            sheet1.col(0).height_mismatch = True
            sheet1.col(0).height = 20*2566
            sheet1.col(2).width = 7000
            sheet1.col(2).height_mismatch = True
            sheet1.col(2).height = 20*256
            sheet1.write(0,2, "رقم السيارة" ,style1)
            sheet1.write(0,0,row[0] ,style1)
            sheet1.write(1,2,"تجديد سنوي" ,style1)
            sheet1.write(1,0,row[1],style1)
            sheet1.write(2,2,"ترخيص سيارات جديدة ",style1)
            sheet1.write(2,0,row[2],style1)
            sheet1.write(3,2,"تجديد تصريح",style1)
            sheet1.write(3,0,row[3],style1)
            sheet1.write(4,2,"استعلام عن مخالفات سائق",style1)
            sheet1.write(4,0,row[4],style1)
            sheet1.write(5,2,"مخالفات ثانوية",style1)
            sheet1.write(5,0,row[5],style1)
            sheet1.write(6,2,"ختم الايصالات المسحوبة من النيابة",style1)
            sheet1.write(6,0,row[6],style1)
            sheet1.write(7,2,"تغريم رخص مسحوبة",style1)
            sheet1.write(7,0,row[7],style1)
            sheet1.write(8,2,"تغريم مخالفات فورية",style1)
            sheet1.write(8,0,row[8],style1)

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/Licence.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def all_car_licence(self):
        self.cursor.execute("SELECT * From tbl_Re_Licences")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        print(self.result[1])

        wb = Workbook()
        sheet1 = wb.add_sheet('All Cars Licences Report')

        sheet1.write(0,0, "رقم السيارة")
        sheet1.write(0,1,"تجديد سنوي")
        sheet1.write(0,2,"ترخيص سيارة جديدة")
        sheet1.write(0,3,"تجديد تصريح")
        sheet1.write(0,4,"استعلام عن مخالفات سائقي")
        sheet1.write(0,5,"مخالفات سنوية")
        sheet1.write(0,6,"ختم الايصالات المسحوبة من النيابة")
        sheet1.write(0,7,"تغريم رخصة مسحوبة")
        sheet1.write(0,8,"تغريم مخالفات فورية")

        row_number = 1
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/Licence.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def Licence_choice(self):
        if self.radioButton_8.isChecked() == True:
            self.one_car_licence()

        if self.radioButton_7.isChecked() == True:
            self.all_car_licence()
    ##################################################3
    # UNRelational Licence
    def one_car_UN_R_licence(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('sheet 1')
        sql = "SELECT * From tbl_Un_Re_Licenses WHERE Car_Num = ?"
        car_number = int(self.le_search.text())
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            sheet1.col(0).width = 7000
            sheet1.col(0).height_mismatch = True
            sheet1.col(0).height = 20*2566
            sheet1.col(2).width = 7000
            sheet1.col(2).height_mismatch = True
            sheet1.col(2).height = 20*256
            sheet1.write(0,2, "رقم السيارة" ,style1)
            sheet1.write(0,0,row[0] ,style1)
            sheet1.write(1,2,"تجديد سنوي" ,style1)
            sheet1.write(1,0,row[1],style1)
            sheet1.write(2,2,"شهادات بيانات - تامينات وضرائب",style1)
            sheet1.write(2,0,row[2],style1)
            sheet1.write(3,2,"ضرائب الاعلانات",style1)
            sheet1.write(3,0,row[3],style1)
            sheet1.write(4,2,"خطابات الصحة والمرور",style1)
            sheet1.write(4,0,row[4],style1)
            sheet1.write(5,2,"بدلات انتقال",style1)
            sheet1.write(5,0,row[5],style1)

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/Licence.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def all_car_UN_R_licence(self):
        self.cursor.execute("SELECT * From tbl_Un_Re_Licenses")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        print(self.result[1])

        wb = Workbook()
        sheet1 = wb.add_sheet('All Cars Licences Report')
        sheet1.write(0,0, "رقم السيارة" )
        sheet1.write(0,1,"تجديد سنوي" )
        sheet1.write(0,2,"شهادات بيانات - تامينات وضرائب")
        sheet1.write(0,3,"ضرائب الاعلانات")
        sheet1.write(0,4,"خطابات الصحة والمرور")
        sheet1.write(0,5,"بدلات انتقال")

        row_number = 1
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/Licence.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def UN_R_Licence_choice(self):
        if self.radioButton_8.isChecked() == True:
            self.one_car_UN_R_licence()

        if self.radioButton_7.isChecked() == True:
            self.all_car_UN_R_licence()


    ##################################################
    def one_car_revenu_report(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('sheet1')
        sql = "SELECT * From tbl_Revnue WHERE Car_Num = ?"
        car_number = str(self.le_search.text())
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')
        sheet1.col(0).width = 7000
        sheet1.col(0).height_mismatch = True
        sheet1.col(0).height = 20*2566
        sheet1.col(2).width = 7000
        sheet1.col(2).height_mismatch = True
        sheet1.col(2).height = 20*256

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            print(row[0])

            sheet1.write(0,2, "رقم السيارة" ,style1)
            sheet1.write(0,0,row[1] ,style1)

            sheet1.write(1,2, "بيع الكهنة" ,style1)
            sheet1.write(1,0,row[2] ,style1)
            sheet1.write(2,2, "تعويضات التامين" ,style1)
            sheet1.write(2,0,row[2] ,style1)
            sheet1.write(3,2, "تعويضات حوادث طرق متعهدين" ,style1)
            sheet1.write(3  ,0,row[2] ,style1)

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/revenu.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def all_car_revenu_report(self):
        self.cursor.execute("SELECT * From tbl_Revnue")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        wb = Workbook()
        sheet1 = wb.add_sheet('All Cars revenu details')

        sheet1.write(0,0,"رقم السيارة")
        sheet1.write(0,1,"بيع الكهنة")
        sheet1.write(0,2,"تعويضات التامين")
        sheet1.write(0,3,"تعويضات حوادث طرف متعهدين")

        row_number = 1
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/revenu.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')

    def Revenu_choice(self):
        if self.radioButton_10.isChecked() == True:
            self.one_car_revenu_report()

        if self.radioButton_9.isChecked() == True:
            self.all_car_revenu_report()
########################################
    def one_car_rents_report(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('sheet1')
        sql = "SELECT * From tbl_Rents WHERE Car_Num = ?"
        car_number = str(self.le_search.text())
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')
        sheet1.col(0).width = 7000
        sheet1.col(0).height_mismatch = True
        sheet1.col(0).height = 20*2566
        sheet1.col(2).width = 7000
        sheet1.col(2).height_mismatch = True
        sheet1.col(2).height = 20*256

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            print(row[0])

            sheet1.write(0,2, "رقم السيارة" ,style1)
            sheet1.write(0,0,row[1] ,style1)

            sheet1.write(1,2, "ايجار شهري " ,style1)
            sheet1.write(1,0,row[2] ,style1)
            sheet1.write(2,2, "غفرة العرب" ,style1)
            sheet1.write(2,0,row[2] ,style1)


        fileName = QFileDialog.getSaveFileName(self, "Save","E:/rents.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def all_car_rents_report(self):
        self.cursor.execute("SELECT * From tbl_Rents")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        wb = Workbook()
        sheet1 = wb.add_sheet('All Cars Rents details')

        sheet1.write(0,0,"رقم السيارة")
        sheet1.write(0,1,"ايجار شهري ")
        sheet1.write(0,2,"غفرة العرب")

        row_number = 1
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/rents.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def Rents_choice(self):
        if self.radioButton_11.isChecked() == True:
            self.one_car_rents_report()

        if self.radioButton_12.isChecked() == True:
            self.all_car_rents_report()
###################################
    def one_car_Clarkat_report(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('sheet1')
        sql = "SELECT * From tbl_Clarkat WHERE Car_Num = ?"
        car_number = str(self.le_search.text())
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')
        sheet1.col(0).width = 7000
        sheet1.col(0).height_mismatch = True
        sheet1.col(0).height = 20*2566
        sheet1.col(2).width = 7000
        sheet1.col(2).height_mismatch = True
        sheet1.col(2).height = 20*256

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            print(row[0])

            sheet1.write(0,2, "رقم السيارة" ,style1)
            sheet1.write(0,0,row[1] ,style1)

            sheet1.write(1,2, "ايجارات كلاركات" ,style1)
            sheet1.write(1,0,row[2] ,style1)
            sheet1.write(2,2, "صيانة كلاركات" ,style1)
            sheet1.write(2,0,row[2] ,style1)

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/Clarkat.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def all_car_Clarkat_report(self):
        self.cursor.execute("SELECT * From tbl_Clarkat")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        wb = Workbook()
        sheet1 = wb.add_sheet('All Cars Rents details')

        sheet1.write(0,0,"رقم السيارة")
        sheet1.write(0,1,"ايجارات كلاركات")
        sheet1.write(0,2,"صيانة كلاركات")

        row_number = 1
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/Clarkat.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def Clarkat_choice(self):
        if self.radioButton_13.isChecked() == True:
            self.one_car_Clarkat_report()

        if self.radioButton_14.isChecked() == True:
            self.all_car_Clarkat_report()
#############################3################################
        #سيارات النقل
###############################################################
    def one_car_Water_Electricty_report(self):
        wb = Workbook()
        sheet1 = wb.add_sheet('sheet1')
        sql = "SELECT * From tbl_Water_Electricty WHERE Car_Num = ?"
        car_number = str(self.le_search.text())
        style1 = easyxf('pattern: pattern solid , fore_color yellow;')
        sheet1.col(0).width = 7000
        sheet1.col(0).height_mismatch = True
        sheet1.col(0).height = 20*2566
        sheet1.col(2).width = 7000
        sheet1.col(2).height_mismatch = True
        sheet1.col(2).height = 20*256

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            print(row[0])

            sheet1.write(0,2, "رقم السيارة" ,style1)
            sheet1.write(0,0,row[1] ,style1)

            sheet1.write(1,2, "امصروفات الكهرباء" ,style1)
            sheet1.write(1,0,row[2] ,style1)
            sheet1.write(2,2, "صمصروفات المياة" ,style1)
            sheet1.write(2,0,row[2] ,style1)

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/WE.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def all_car_Water_Electricty_report(self):
        self.cursor.execute("SELECT * From tbl_Water_Electricty")
        self.result = self.cursor.fetchall()
        print(self.result[0])
        wb = Workbook()
        sheet1 = wb.add_sheet('Water_Electricty')

        sheet1.write(0,0,"رقم السيارة")
        sheet1.write(0,0,"مصروفات الكهرباء")
        sheet1.write(0,0,"مصروفات المياة")

        row_number = 1
        for row in self.result :
            column_num = 0
            for item in row :
                sheet1.write(row_number,column_num,str(item) )
                column_num = column_num + 1
            row_number = row_number + 1

        fileName = QFileDialog.getSaveFileName(self, "Save","E:/WE.xls","Excel (.xls )")
        wb.save(fileName)
        self.statusbar.showMessage('تم طباعة التقرير بنجاح')


    def Water_Electricty_choice(self):
        if self.radioButton_17.isChecked() == True:
            self.one_car_Water_Electricty_report()

        if self.radioButton_18.isChecked() == True:
            self.all_car_Water_Electricty_report()
##########################################################################################
##########################################################################################
        ## Delete Data
    def delete_car(self):
        self.cursor.execute("DELETE FROM tbl_Car WHERE Car_Num ='%s'"%int(self.le_search.text()))
        self.cursor.execute("DELETE FROM tbl_Clarkat WHERE Car_Num ='%s'"%int(self.le_search.text()))
        self.cursor.execute("DELETE FROM tbl_Fuel WHERE Car_Num ='%s'"%int(self.le_search.text()))
        self.cursor.execute("DELETE FROM tbl_R_Maintanance WHERE Car_Num ='%s'"%int(self.le_search.text()))
        self.cursor.execute("DELETE FROM tbl_Re_Licences WHERE Car_Num ='%s'"%int(self.le_search.text()))
        self.cursor.execute("DELETE FROM tbl_Rents WHERE Car_Num ='%s'"%int(self.le_search.text()))
        self.cursor.execute("DELETE FROM tbl_Revnue WHERE Car_Num ='%s'"%int(self.le_search.text()))
        self.cursor.execute("DELETE FROM tbl_Water_Electricty WHERE Car_Num ='%s'"%int(self.le_search.text()))
        self.cursor.commit()
        self.label_14.setText('تم مسح المعلومات بنجاح')
        self.statusbar.showMessage('تم مسح المعلومات بنجاح')
        print('done')

##########################################################################################
##########################################################################################
    # Show One Result Fields
    def Retrieve_Car_Data(self):
        sql = "SELECT * From tbl_Car WHERE Car_Num = ?"
        car_number = int(self.le_search.text())

        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            print(row[0])
            print(row[7])
            self.le_car_number_2.setText(row[1])
            self.le_owner_company_2.setText(row[2])
            self.le_branch_2.setText(row[3])

            if row[4] == 'بيع' :
                self.cm_service_mode_2.setCurrentIndex(0)

            elif row[4] == 'ملاكي':
                self.cm_service_mode_2.setCurrentIndex(1)

            elif row[4] == 'امداد':
                self.cm_service_mode_2.setCurrentIndex(2)

            else :
                self.cm_service_mode_2.setCurrentIndex(3)

            self.le_shaceh_number_2.setText(row[5])
            self.le_motor_number_2.setText(row[6])

            if row[7] == 'سولار' :
                self.cm_full_kind_2.setCurrentIndex(0)
            else:
                self.cm_full_kind_2.setCurrentIndex(1)

            self.le_car_kind_2.setText(row[8])
            self.le_car_model_2.setText(str(row[9]))
            self.le_car_load_2.setText(str(row[10]))
            self.le_car_weight_2.setText(str(row[11]))
            self.le_shape_2.setText(row[12])
            self.le_color_2.setText(row[13])



    def Retrieve_Fuel_Data(self):
        sql = "SELECT * From tbl_Fuel WHERE Car_Num = ?"
        car_number = int(self.le_search.text())
        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            self.le_leters_number_2.setText(str(row[2]))
            self.le_price_2.setText(str(row[3]))
            self.le_meter_number_2.setText(str(row[4]))
            self.lcd_number_2.display(row[5])
            self.le_tips_2.setText(str(row[6]))


    def Retrieve_R_Maintenance_Data(self):
        sql = "SELECT * From tbl_R_Maintanance WHERE Car_Num = ?"
        car_number = int(self.le_search.text())
        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            self.lineEdit_9.setText(row[1])
            self.lineEdit_33.setText(row[2])
            self.lineEdit_15.setText(row[3])
            self.lineEdit_29.setText(row[4])
            self.lineEdit_30.setText(row[5])
            self.lineEdit_31.setText(row[6])
            self.lineEdit_35.setText(row[7])
            self.lineEdit_34.setText(row[8])



    def Retrieve_Un_R_Maintenance_Data(self):
        sql = "SELECT * From tbl_Un_Re_Maintainance WHERE Form = ? AND [To]= ?"
        max_value = self.dateEdit_6.text()
        minvalue = self.dateEdit_7.text()
        for row in self.cursor.execute(sql,(max_value,minvalue)):
            print(str(row))
            self.le_value_2.setText(str(row[3]))


    def Retrieve_R_Licence_Data(self):
        sql = "SELECT * From tbl_Re_Licences WHERE Car_Num = ?"
        car_number = int(self.le_search.text())
        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            self.lineEdit_16.setText(row[1])
            self.lineEdit_18.setText(row[2])
            self.lineEdit_17.setText(row[3])
            self.lineEdit_20.setText(row[4])
            self.lineEdit_19.setText(row[5])
            self.lineEdit_27.setText(row[6])
            self.lineEdit_21.setText(row[7])
            self.lineEdit_28.setText(row[8])


    def Retrieve_Un_R_Licence_Data(self):
        sql = "SELECT * From tbl_Un_Re_Licenses WHERE Car_Num = ?"
        car_number = int(self.le_search.text())
        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            self.le_licenses_value_4.setText(row[1])
            self.le_licenses_value_5.setText(row[2])
            self.le_licenses_value_7.setText(row[3])
            self.le_licenses_value_6.setText(row[4])


    def Retrieve_Revenu_Data(self):
        sql = "SELECT * From tbl_Revnue WHERE Car_Num = ?"
        car_number = str(self.le_search.text())
        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            self.le_revenues_value_2.seText(row[1])
            self.le_revenues_value_3.seText(row[2])
            self.le_revenues_value_4.seText(row[3])



    def Retrieve_Rents_Data(self):
        sql = "SELECT * From tbl_Rents WHERE Car_Num = ?"
        car_number = str(self.le_search.text())
        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            self.le_rents_value_3.setText(row[1])
            self.le_rents_value_2.setText(row[2])


    def Retrieve_Clarkat_Data(self):
        sql = "SELECT * From tbl_Clarkat WHERE Car_Num = ?"
        car_number = str(self.le_search.text())
        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            self.le_kind_value_3.setText(row[1])
            self.le_kind_value_2.setText(row[2])


    def Retrieve_Cars_Data(self):
        pass


    def Retrieve_Weter_Electricty_Data(self):
        sql = "SELECT * From tbl_Water_Electricty WHERE Car_Num = ?"
        car_number = str(self.le_search.text())
        for row in self.cursor.execute(sql ,[(car_number)]):
            print(str(row))
            self.le_elctric_value_2.setText(row[1])
            self.le_water_water_value_2.setText(row[2])

##########################################################################################
##########################################################################################
    ## USERS
    def add_user(self):
        self.cursor.execute("INSERT INTO tbl_Users(User_Name,Password,Add_Car,Licence,Rents,Fuel,Revenu,Clarkat,Maintenance,Water_Electric,Report,Users)"
                            " VALUES (?,?,?,?,?,?,?,?,?,?,?,?)" ,(self.lineEdit_8.text(),self.lineEdit_2.text(),0,0,0,0,0,0,0,0,0,0))
        self.cursor.commit()
        self.statusbar.showMessage("تمت اضافة المستخدم بنجاح")
        print('Done')

    def add_user_privelige(self):
        if self.checkBox.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Add_Car = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_4.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Licence = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_7.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Rents = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_2.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Fuel = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_5.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Revenu = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_8.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Clarkat = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_3.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Maintenance = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_9.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Water_Electric = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_11.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Report = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_10.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Cars_Rents = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        if self.checkBox_12.isChecked() == True :
            self.cursor.execute("UPDATE tbl_Users SET Users = 1 WHERE User_Name = ? AND Password = ? " ,
                                (self.lineEdit_8.text() ,self.lineEdit_2.text()))

        self.cursor.commit()
        self.statusbar.showMessage('تم اضافة الصلاحيات بنجاح')
    ##############################################################
    ##############################################################
    def update_user_password(self):
        self.cursor.execute("UPDATE tbl_Users SET Password = ?  WHERE User_Name = ? " ,(self.lineEdit_14.text() , self.lineEdit_51.text()))
        self.cursor.commit()
        self.statusbar.showMessage('تم تعديل الباسورد ب نجاح')


    def update_user_privilage(self):
        pass


###########################################################################################
    def btnstate(self,r):
         if r.isChecked() == True:
            print (" 1 is selected")  # Execute something
         elif r.isChecked() == False:
            print (" 1 is deselected")  #execute something



def main_app():
    app = QApplication(sys.argv)
    #window = login()
    window = main() 
    window.show()
    app.exec_()

if __name__ == "__main__":
    main_app()
