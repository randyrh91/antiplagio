#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# https://pythones.net

import webbrowser
import fitz
import os
import win32com.client
import docx
import re
from difflib import SequenceMatcher as SM
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5 import QtWidgets
from string import punctuation
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.feature_extraction.text import TfidfVectorizer

language_stopwords = stopwords.words('russian')
non_words = list(punctuation)

qtCreatorFile = "antiplagio.ui"

Ui_MainWindow, QtBaseClass = uic.loadUiType(
    qtCreatorFile)


class VentanaPrincipal(QtWidgets.QMainWindow, Ui_MainWindow):

    error = False

    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.pushButtonAddFile.clicked.connect(self.add_files)
        self.pushButton_DeleteFile.clicked.connect(self.delete_file)
        self.pushButtonExecute.clicked.connect(self.execute)
        self.pushButtonAddFileOriginal.clicked.connect(self.add_original_file)

    def add_files(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл", "", "docx (*.docx);;pdf (*.pdf);;doc (*.doc)", options=options)
        if fileName:
            self.listWidget.addItem(fileName)

    def add_original_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл", "", "docx (*.docx);;pdf (*.pdf);;doc (*.doc)", options=options)
        if fileName:
            self.fileWidget.addItem(fileName)

    def search_similarity(self):
        fileList = [self.listWidget.item(i).text() for i in range(self.listWidget.count())]
        text_read_file = []        
        all = []
        max = []
        files = []
        if self.fileWidget.count() == 1:
            ruta_orig_file = self.fileWidget.item(0).text()
            files = [ruta_orig_file] + fileList
            text_original = self.process_file(ruta_orig_file)
            text_read_file.append(text_original)
            for read_file in fileList:
                text = self.process_file(read_file)
                text_read_file.append(text)
            if(len(text_read_file) > 1):
                vectorizer = TfidfVectorizer()
                X = vectorizer.fit_transform(text_read_file)
                similarity_matrix = cosine_similarity(X, X)
                all, max = self.search_max(similarity_matrix)
            else:
                self.labelExecute.setText(
                    "в список необходимо добавить более одного документа.")
                self.labelExecute.setStyleSheet("background-color: lightsalmon")
        else:
            self.labelExecute.setText(
                    "вы должны добавить документ")
            self.labelExecute.setStyleSheet("background-color: lightsalmon")
        return [all, max, files]

    def delete_file(self):
        for SelectedItem in self.listWidget.selectedItems():
            self.listWidget.takeItem(
                self.listWidget.indexFromItem(SelectedItem).row())

    def similarity(self, all, max, files):
        
        text_1 = self.process_file(files[0])
        text_2 = self.process_file(files[max[0]])
        par_1 = self.parag_div(text_1)
        par_2 = self.parag_div(text_2)
        total = 0
        sim = 0
        text_analiz = ""
        for p1 in par_1:
            if(p1.strip() != ''):
                sent_1 = self.sent_div(p1)
                text_analiz += "<p>"
                for s1 in sent_1:       
                    porcSim = 0
                    if(s1.strip() != ''):
                        total +=1
                        isText = False
                        for p2 in par_2:
                            if(p2.strip() != ''):
                                sent_2 = self.sent_div(p2)
                                for s2 in sent_2:
                                    if(s2.strip() != ''):
                                        try:
                                            s1 = self.clean_text(s1)
                                            s2 = self.clean_text(s2)
                                            sim_val = SM(None, s1.strip(), s2.strip()).ratio()
                                            if(sim_val > 0.75):
                                                isText = True
                                                porcSim = sim_val
                                        except ValueError:
                                            continue
                        if(isText):
                            sim += 1 
                            text_analiz += "<span style = 'color: #64AA26' title='Сходство ("+ str(round(porcSim*100,2)) +"%)'>"+ s1 +"</span>"
                        else:
                            text_analiz += "<span>"+ s1 +"</span>"
                text_analiz += "</p>" 
        
        porc_sim = round((sim * 100)/total, 2)
        
        html = "<!DOCTYPE html <html lang='ru'> <head><style>td {border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: Gray;}.bar-out {background-color: white;border: 1px solid gray;border-width: 0 0 1px 1px;}.bar-in {width: 100%;height: 6px;}"
        
        html += "h2 {font-size: medium;color: #336699;}h1 {font-size: 24px;color: #336699;}.fontemedia {font-size: 16px;margin: 0px;font-weight: bold;padding: 0px;}.fontegrande {font-size: 20px;font-weight: bold;margin: 0px;padding: 0px;color: blue;}.box {background-color: White;border-style: none;border-color: Gray;padding: 5px;}</style>"
        
        # agregar estilos
        
        html += "</head><body> <h1>Результаты анализа</h1><blockquote><div><span style='font-size: 14px;'><div style='padding-bottom: 7px;'><strong>Проанализированный файл:</strong><div><a href='"+files[0]+"' target='_blank'>" + files[0]+ "</a></div></div><div style='padding-bottom: 7px;'><strong>Более похожий файл:</strong><div><a href='"+files[max[0]]+"' target='_blank'>" + files[max[0]]+ "</a></div></div></div></span></div></blockquote><h2>Статистика</h2><blockquote>"

        html += "<div class='' style=''><p class='fontemedia'>Процентов сходства: <span class='fontegrande'>"+ str(porc_sim) + "%</span></p><p>Процент текста с похожими выражениями, найденными в файле наибольшего сходства.</p><div><div class='bar-out' style=''><div class='bar-in' style='background-color: #529756; width:"+ str(porc_sim) +"%;'></div></div></div></div></blockquote>"
        
        html+="<div><h2>Все проанализированные файлы</h2><blockquote>"
        
        #Aqui la tabla

        html += "<table cellpadding='3px' cellspacing='0px'><tbody><tr><td><strong>Адрес файла</strong></td><td><strong>Подобие</strong></td></tr>"
       
        for i in range(len(files)):
            if i!=0:
                html+= "<tr><td><a href='"+files[i]+"' target='_blank'>"+files[i]+"</a></td><td>"+str(all[i])+"%</td></tr>"
        
        html += "</tbody></table>"
        
        html += "</div><div><h2>Анализируемая текст:</h2><blockquote><div class = 'box'>"

        html += text_analiz

        html += "<div></blockquote></div></body><footer></footer></html>"
        
        self.generate_html(html)

    def execute(self):
        self.labelExecute.setText("Выполняя процесс...")
        self.labelExecute.setStyleSheet("background-color: lightgreen")
        if(self.error == False):
            all, max, files = self.search_similarity()
            self.similarity(all, max, files)
            self.labelExecute.setText("Выполнение завершено")
            self.labelExecute.setStyleSheet("background-color: lightgreen")
        else:
            self.error = False

    def parag_div(self, text):
        return re.split("(\\t|\\r|\\f|\\v|\\n)+(?![а-яё])", text)

    def sent_div(self, text):
        return re.split('(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', text)

    def clean_text(self, text):
        text = text.replace("\n","")
        text = text.replace("\t","")
        text = text.replace("\r","")
        text = text.replace("\v","")
        text = text.replace("\f","")
        return text

    def remove_stop_words(self, text):
        tokens = word_tokenize(text)
        clean_tokens = tokens[:]
        for token in tokens:
            if token in language_stopwords:
                clean_tokens.remove(token)
        return (" ").join(clean_tokens)

    def getTextPDF(self, doc_pdf):
        text = ''
        for page in doc_pdf:
            text += page.getText()
        return text

    def getDataPDF(self, file_name):
        page_content = ""
        try:
            pdf_file = fitz.open(file_name)
            page_content = self.getTextPDF(pdf_file).replace('-\n', '')
        except Exception:
            self.labelExecute.setText("Ошибка при попытке прочитать документ.")
            self.labelExecute.setStyleSheet("background-color: lightsalmon")
            self.error = True
        return page_content

    def getDataDocx(self, file_name):
        try:
            doc = docx.Document(file_name)
            fullText = []
            for para in doc.paragraphs:
                fullText.append(para.text)
            fullText = '\n'.join(fullText)
        except Exception:
            self.labelExecute.setText("Ошибка при попытке прочитать документ.")
            self.labelExecute.setStyleSheet("background-color: lightsalmon")
            self.error = True
        return fullText

    def getDataDoc(self, file_name):
        fullText = ""
        try:
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(file_name, False, False, False)
            fullText = doc.Range().Text
            doc.Close()
            word.Quit()
        except Exception:
            self.labelExecute.setText("Ошибка при попытке прочитать документ.")
            self.labelExecute.setStyleSheet("background-color: lightsalmon")
            self.error = True
        return fullText

    def generate_html(self, html):
        f = open('report/reporte.html', 'w', encoding =  'utf-8')
        f.write(html)
        f.close()
        webbrowser.open_new_tab('report/reporte.html')

    def process_file(self, file_name):
        doc, ext = os.path.splitext(file_name)
        if ext == ".doc":
            file_content = self.getDataDoc(file_name)
        elif ext == ".docx":
            file_content = self.getDataDocx(file_name)
        else:
            file_content = self.getDataPDF(file_name)
        return file_content

    def search_max(self, matrix):
        max = [0 , 0]
        all = []
        max_elem = 0
        for j in range(len(matrix[0])):
            elem = round(matrix[j][0]*100, 2)
            all.append(elem)
            if(elem > max_elem and 0 != j):
                max_elem = elem
                max[0] = j
                max[1] = elem
        return [all, max]

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    ventana = VentanaPrincipal()
    ventana.show()
    app.exec_()
