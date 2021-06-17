from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
import sys
import pickle
import time
from selenium import webdriver
import pandas
import fbchat
import time
import re

ui, _ = loadUiType("D:/backup/Desktop/DeskFiles/python/fb.ui")
# groupLink = "https://www.facebook.com/groups/1714991985474724/"

{'status': 'false',
 'id': ['100001866524038', '100039403572931', '100035315856110', '2501351436755956', '2348600322130866', '2266254706829447', '100047455298727', '100055108455757', '100050454609552', '100001302391682', '100053868287213', '1561089551', '100049249912811', '100002511920671', '100028097502889', '100001676475705', '100054580747558', '100001468294864', '100027596496542', '100007345878640', '100001247891139', '100007973192242', '100010675146368', '100026919743963', '100002344343295', '100041529697788', '100042136448395', '1539780132', '100003909166783', '100012283508762', '100003846242842', '100002293569410', '100026159365818', '100035356086304', '100041443491786', '100002606427455'],
 'url': ['https://www.facebook.com/amro1111', 'https://www.facebook.com/profile.php?id=100039403572931', 'https://www.facebook.com/profile.php?id=100035315856110', 'https://www.facebook.com/mohammed.elbabarawy81/', 'https://www.facebook.com/العيون-الجميلة-للنظارات-2348600322130866/', 'https://www.facebook.com/YasserEdri/', 'https://www.facebook.com/profile.php?id=100047455298727', 'https://www.facebook.com/profile.php?id=100055108455757', 'https://www.facebook.com/profile.php?id=100050454609552', 'https://www.facebook.com/ahmed.elsayed.58910049', 'https://www.facebook.com/mahmoud.alahlawy.779', 'https://www.facebook.com/adel.haggag', 'https://www.facebook.com/kamel.abdelalim.1', 'https://www.facebook.com/karm.opticshorghada', 'https://www.facebook.com/profile.php?id=100028097502889', 'https://www.facebook.com/moshera.mustafa.1', 'https://www.facebook.com/profile.php?id=100054580747558', 'https://www.facebook.com/Ashraf.Hamdy.Abd.Elaziz', 'https://www.facebook.com/profile.php?id=100027596496542', 'https://www.facebook.com/profile.php?id=100007345878640', 'https://www.facebook.com/sona.mohamed.754', 'https://www.facebook.com/hamam.said.5', 'https://www.facebook.com/profile.php?id=100010675146368', 'https://www.facebook.com/afifa.kaid', 'https://www.facebook.com/magdy.shafek.58', 'https://www.facebook.com/suzaz.ghb', 'https://www.facebook.com/yasser.kassar.1', 'https://www.facebook.com/ahmed.m.abbas.75', 'https://www.facebook.com/aona.mostafa', 'https://www.facebook.com/ryan.shi.714', 'https://www.facebook.com/medo.tiger.923', 'https://www.facebook.com/adel.nasr.73', 'https://www.facebook.com/profile.php?id=100026159365818', 'https://www.facebook.com/malak.elsokary.98', 'https://www.facebook.com/profile.php?id=100041443491786', 'https://www.facebook.com/farahbahaaa'],
 'is_viewer_friend': [False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False],
 'gender': ['Female', 'Female', 'male', 'Female', 'male', 'Female', 'Female', 'Female', 'Female', 'Female', 'Female', 'Female', 'male', 'Female', 'Female', 'Female', 'Female', 'Female', 'Female', 'Female', 'male', 'Female', 'Female', 'Female', 'Female', 'male', 'Female', 'Female', 'Female', 'Female', 'male', 'Female', 'male'],
 'name': ['Amr Ragab', 'Omar Ahmed', 'انچى سالم', 'Nazra Opticsنظره للبصريات', 'العيون الجميلة للنظارات', 'Brand is less and sweat heart', 'مهدي الطالبي', '齐晶', 'عہسہليہه عہسہل', 'Ahmed El-Sayed', 'Mahmoud Alahlawy', 'Adel Mahmoud Haggag', 'Kamel Abdelalim', 'Karm Optics Horghada', 'عباس للمبيعات', 'Moshera Mustafa', 'احمد الهنداوى', 'Ashraf Hamdy', 'خلف صالح', 'محمد أيمن الترامسى', 'Ahmed AE', 'Hamam Said', 'محمد النحراوى', 'Aid Salsabil', 'Magdy Shafek', 'Mohamed Shereen', 'Yasser Ali', 'Ahmed Moustafa Abbas', 'Onaa Mostafa', 'Peter Sun', 'Opt Mohamed Elhdad', 'Adel Nasr', 'عبدالله ربيع', 'Malak El Sokary', 'Mohammad Ibrahim', 'Farah Bahaa']
 }


def init_excel_file(filename, *dataA, mode):
    # data_dic={}
    re = pandas.read_excel(filename)
    global column_names
    column_names = []
    for col_name in re.keys():
        column_names.append(col_name)
    if mode == 'read':
        data = re.to_dict()
        return data
    elif mode == 'write':
        data = dataA[0]
        writer = pandas.ExcelWriter(filename, engine='xlsxwriter')
        df = pandas.DataFrame(data)
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()
    elif mode == 'append':
        old_data = pandas.read_excel(filename)
        new_data = pandas.DataFrame(dataA[0])
        print(new_data)
        all_data = pandas.concat([old_data, new_data], ignore_index=1)
        print(all_data)
        writer1 = pandas.ExcelWriter(filename, engine='xlsxwriter')
        # df1 = pandas.DataFrame(all_data)
        all_data.to_excel(writer1, sheet_name='Sheet1', index=False)
        writer1.save()


class mainAPP(QMainWindow, ui):

    def __init__(self, parent=None):
        super(mainAPP, self).__init__(parent)
        QMainWindow.__init__(self)
        self.id_dic = {}
        self.setupUi(self)
        self.handelButtons()
        self.placholder()
        self.tt.hide()
        self.same.toggled.connect(lambda: self.placholder(self.same))
        self.one.toggled.connect(lambda: self.placholder(self.one))

    def handelButtons(self):
        self.loadbtn.clicked.connect(self.loadData)
        self.openfb.clicked.connect(self.openBrowser)
        self.loadcookies.clicked.connect(self.get_and_save_cookies)
        self.opeGlink.clicked.connect(self.openGroupLink)
        self.scrollbtn.clicked.connect(self.scrolldown)
        # self.addbtn.clicked.connect(self.addD)
        self.sendbtn.clicked.connect(self.read_exported_file)
        self.exportbtn.clicked.connect(self.export_to_excel)
        self.selectfile.clicked.connect(self.Browse_files)
        self.openbtn.clicked.connect(self.open_account)

    def openBrowser(self):
        global webBrowser
        webBrowser = webdriver.Chrome(
            "D:/backup/Desktop/DeskFiles/python/chromedriver.exe")
        url = "https://www.facebook.com"
        webBrowser.get(url)

    def get_and_save_cookies(self):
        print("enter")
        # save current cookie in pickle file
        pickle.dump(webBrowser.get_cookies(), open(
            "D:/backup/Desktop/DeskFiles/python/cookies2.pkl", "ab+"))
        # load pickle file
        cookies = pickle.load(
            open("D:/backup/Desktop/DeskFiles/python/cookies2.pkl", "rb"))
        print("SD")
        # start add cookies to borwser
        for cookie in cookies:
            webBrowser.add_cookie(cookie)
        print("cookie loaded successfully !")

    def openGroupLink(self):
        G_url = self.groupurl.text()
        webBrowser.get(G_url)

    #     https://www.facebook.com/messages/t/

    def loadData(self):
        wanted = ['id', 'url', 'name', 'gender', 'is_viewer_friend']
        print("start load ids")
        print("web start")
        global ids
        ids = []
        members = webBrowser.find_elements_by_css_selector(
            "div[class='clearfix _60rh _gse'][data-name='GroupProfileGridItem']")
        print(len(members), "id founded ! ")
        founded = len(members)
        self.lcd1.display(founded)
        for i in members:
            preresult = re.findall("\d", i.get_attribute("data-testid"))
            result = "".join(preresult)
            ids.append(result)
        print(ids)

        print("load ids done !")
        print("start connection to get ids data")
        global username
        username = 'hassan.cool410@yahoo.com'
        global client
        client = fbchat.Client(username, 'hassan69199920201010')
        print("connection done !")
        self.mytable.setRowCount(0)
        added = 0
        for row_count, member in enumerate(ids):
            added += 1
            try:
                mem = client._fetchInfo(member)
            except Exception:
                print(Exception)
            # print(mem)
            self.mytable.insertRow(row_count)
            col_count = 0
            for val in mem.values():
                self.id_dic['status'] = 'false'
                for key, value in val.items():
                    if str(value) == "ThreadType.PAGE":
                        self.id_dic['gender'].append('PAGE')
                        self.id_dic['is_viewer_friend'].append('PAGE')

                    if key in wanted:
                        if key == 'gender':
                            if value == 2:
                                value = 'male'
                            elif value == 1:
                                value = 'Female'
                        if key in self.id_dic:
                            self.id_dic[key].append(value)
                        else:
                            self.id_dic[key] = []
                            self.id_dic[key].append(value)
                        self.mytable.setItem(
                            row_count, col_count, QtWidgets.QTableWidgetItem(str(value)))
                        col_count += 1
                        self.lcd2.display(added)
                        self.lcd3.display(founded-added)
                        QApplication.processEvents()
        QMessageBox.information(
            self, 'Compelete', str(row_count+1)+" added successfuly !")

    def scrolldown(self):
        QMessageBox.information(
            self, 'Wait', "Please wait until Scroll down finished !")
        self.tt.show()
        lenOfPage = webBrowser.execute_script(
            "window.scrollTo(0, document.body.scrollHeight+10000);return document.body.scrollHeight;")
        match = False
        while match == False:
            lastCount = lenOfPage
            time.sleep(3)
            lenOfPage = webBrowser.execute_script(
                "window.scrollTo(0, document.body.scrollHeight+10000);return document.body.scrollHeight;")
            QApplication.processEvents()
            # print("lastCount", lastCount)
            # print("lenOfPage", lenOfPage)
            time.sleep(2)
            if lastCount == lenOfPage:
                match = True
        self.tt.hide()
        QMessageBox.information(
            self, 'Compelete', "Scroll down has been finished !")

    def send_msg(self, mode, data):
        username = 'hassan.cool410@yahoo.com'
        client = fbchat.Client(username, 'hassan69199920201010')
        if self.memnum.text() == "":
            QMessageBox.warning(self, "Attention !",
                                "number of people at least must be 1 person !")
        else:
            numOfPeople = int(self.memnum.text())
            sent = 0
            self.remainsent.display(numOfPeople)
            if self.same.isChecked():

                for i in range(numOfPeople):
                    try:
                        if mode == 'list':
                            client.send(fbchat.models.Message(
                                str(self.msgcontent.toPlainText())), ids[i])
                        elif mode == 'element':
                            client.send(fbchat.models.Message(
                                str(self.msgcontent.toPlainText())), data[0])
                    except Exception as e:
                        print("Exception", e)
                    sent += 1
                    numOfPeople -= 1
                    self.sentnum.display(sent)
                    self.Bar.setValue(int((sent/(sent+numOfPeople))*100))
                    self.remainsent.display(numOfPeople)
                    QApplication.processEvents()
                QMessageBox.information(
                    self, "complete", "messages have been sent successfully !")

            #  diffirent msg
            elif self.one.isChecked():
                s = [str(self.msgcontent.toPlainText())]
                listOfMsg = s[0].split('\n')
                for i, j in zip(range(numOfPeople), listOfMsg):
                    client.send(j, ids[i])
                    sent += 1
                    self.sentnum.display(sent)
                    numOfPeople -= 1
                    self.remainsent.display(numOfPeople)
                    self.Bar.setValue((sent / (sent + numOfPeople)) * 100)
                    QApplication.processEvents()
                QMessageBox.information(
                    self, "complete", "messages have been sent successfully !")

    def placholder(self, *a):
        if self.one.isChecked():
            self.msgcontent.setPlaceholderText(
                "Enter your messages separeted by 'ENTER'")
        if self.same.isChecked():
            self.msgcontent.setPlaceholderText("Type your message")

    def export_to_excel(self):
        print(self.id_dic)
        saveLocation = QFileDialog.getSaveFileName(
            self, caption="save as", directory=".", initialFilter='.xlsx', filter='.xlsx')
        if saveLocation[0] == "":
            return
        else:
            filename = saveLocation[0]+'.xlsx'
            pandas.ExcelWriter(filename, engine='xlsxwriter').save()
            init_excel_file(filename, self.id_dic, mode='append')
        # print("dsa")

    def Browse_files(self):
        filepath = QFileDialog.getOpenFileName(
            self, caption="save as", directory=".",  filter='all files (*.xlsx*)')
        print(filepath[0])
        self.filepath.setText(str(filepath[0]))

    def read_exported_file(self):
        numOfPeople = int(self.memnum.text())
        # username = 'hassan.cool410@yahoo.com'
        dc = init_excel_file(self.filepath.text(), mode='read')
        count = 0
        if self.memnum.text() == "":
            QMessageBox.warning(self, "Attention !",
                                "number of people at least must be 1 person !")
        else:
            # client = fbchat.Client(username, 'hassan69199920201010')
            sent = 0
            self.remainsent.display(numOfPeople)
            if self.same.isChecked():
                # webBrowser = webdriver.Chrome(
                #     "D:/backup/Desktop/DeskFiles/python/chromedriver.exe")
                # url = "https://www.facebook.com/messages/t/"
                # webBrowser.get(url)
                for (key, value) in dc['status'].items():
                    if numOfPeople != 0:
                        # print("enter if")
                        if str(value) == 'False' or str(value) == 'faild':
                            try:
                                print(dc['id'][key])
                                webBrowser3.get(
                                    "https://www.facebook.com/messages/t/"+str(dc['id'][key]))
                                time.sleep(3)
                                webBrowser3.find_element_by_class_name('_1mf._1mk').send_keys(
                                    str(self.msgcontent.toPlainText()))
                                webBrowser3.find_element_by_class_name(
                                    'oajrlxb2.gs1a9yip.g5ia77u1.mtkw9kbi.'
                                    'tlpljxtp.qensuy8j.ppp5ayq2.goun2846.'
                                    'ccm00jje.s44p3ltw.mk2mc5f4.rt8b4zig.'
                                    'n8ej3o3l.agehan2d.sk4xxmp2.rq0escxv.'
                                    'nhd2j8a9.pq6dq46d.mg4g778l.btwxx1t3.'
                                    'pfnyh3mw.p7hjln8o.knvmm38d.cgat1ltu.'
                                    'bi6gxh9e.kkf49tns.tgvbjcpo.hpfvmrgz.'
                                    'cxgpxx05.dflh9lhu.sj5x9vvc.scb9dxdr.'
                                    'l9j0dhe7.i1ao9s8h.esuyzwwr.f1sip0of.'
                                    'du4w35lb.lzcic4wl.abiwlrkh.p8dawk7l').click()
                                dc['status'][key] = 'Sent'
                                init_excel_file(
                                    self.filepath.text(), dc, mode='write')
                                time.sleep(3)

                                # client.send(fbchat.models.Message(
                                # str(self.msgcontent.toPlainText())), dc['id'][key])
                            except Exception:
                                print("can't send !")
                                dc['status'][key] = 'faild'
                            count += 1
                            sent += 1
                            numOfPeople -= 1
                            self.sentnum.display(sent)
                            self.Bar.setValue(
                                int((sent / (sent + numOfPeople)) * 100))
                            self.remainsent.display(numOfPeople)
                            QApplication.processEvents()
                        else:
                            continue
                QMessageBox.information(
                    self, "complete", "messages have been sent successfully !")
            elif self.one.isChecked():
                s = [str(self.msgcontent.toPlainText())]
                listOfMsg = s[0].split('\n')
                for key, value, j in zip(dc['status'].items(), listOfMsg):
                    if count <= numOfPeople:
                        if str(value) == 'False':
                            client.send(j, dc['id'][key])
                            dc['status'][key] == 'Sent'
                            count += 1
                            sent += 1
                            self.sentnum.display(sent)
                            numOfPeople -= 1
                            self.remainsent.display(numOfPeople)
                            self.Bar.setValue(
                                (sent / (sent + numOfPeople)) * 100)
                            QApplication.processEvents()
                        else:
                            continue
                QMessageBox.information(
                    self, "complete", "messages have been sent successfully !")
        init_excel_file(self.filepath.text(), dc, mode='write')

    def open_account(self):
        global webBrowser3
        webBrowser3 = webdriver.Chrome(
            "D:/backup/Desktop/DeskFiles/python/chromedriver.exe")
        url3 = "https://www.facebook.com"
        webBrowser3.get(url3)
        input("ddddddd")
        # ime.sleep(2)


def main():
    app = QApplication(sys.argv)
    window = mainAPP()
    window.show()
    app.exec_()


if __name__ == "__main__":
    main()
