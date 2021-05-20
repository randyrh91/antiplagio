#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# https://pythones.net

import sys
from PyPDF2 import PdfFileWriter, PdfFileReader
from io import StringIO


from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, inch, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.graphics.charts.legends import Legend
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.shapes import Drawing, String
from reportlab.platypus import Paragraph
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.validators import Auto

import plotly
from plotly.offline import iplot, init_notebook_mode
import plotly.graph_objs as go

from fpdf import FPDF

import fitz
from docx2pdf import convert
import pypandoc
import os
import win32com.client
import docx
import re
import docx2txt
from time import time
from docx import Document
from tidylib import tidy_document
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5 import QtCore, QtGui, QtWidgets
from string import punctuation
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize
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
        self.pushButtonAddFile.clicked.connect(self.add_file)
        self.pushButton_DeleteFile.clicked.connect(self.delete_file)
        self.pushButtonExecute.clicked.connect(self.execute)

    def add_file(self):
        eventObj = self.sender().objectName()
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл", "", "pdf (*.pdf);;docx (*.docx);;doc (*.doc)", options=options)
        if fileName:
            self.listWidget.addItem(fileName)

    def search_similarity(self):
        readedFileList = [self.listWidget.item(i).text() for i in range(self.listWidget.count())]
        text_read_file = []
        read_rute_file = []
        porcent = 0
        for read_file in readedFileList:
            text = self.process_file(read_file, '0')
            text_read_file.append(self.clean_text(text[0]))
        if(len(text_read_file) > 1):
            vectorizer = TfidfVectorizer()
            X = vectorizer.fit_transform(text_read_file)
            similarity_matrix = cosine_similarity(X, X)
            index_file1, index_file2, porcent = self.search_max(similarity_matrix)
            read_rute_file.append(readedFileList[index_file1])
            read_rute_file.append(readedFileList[index_file2])
        else:
            self.labelExecute.setText(
                "в список необходимо добавить более одного документа.")
            self.labelExecute.setStyleSheet("background-color: lightsalmon")
        return [read_rute_file, porcent]

    def highlightPDF(self, doc, text):
        inst_counter = 0
        for pi in range(doc.pageCount):
            page = doc[pi]
            text_instances = page.searchFor(text)
            five_percent_height = (page.rect.br.y - page.rect.tl.y)*0.05
            for inst in text_instances:
                inst_counter += 1
                highlight = page.addHighlightAnnot(inst)
            doc.saveIncr()

    def delete_file(self):
        for SelectedItem in self.listWidget.selectedItems():
            self.listWidget.takeItem(
                self.listWidget.indexFromItem(SelectedItem).row())

    def similarity(self, ruteFile1, ruteFile2):
        doc_1, text_1 = self.process_file(ruteFile1, '1')
        doc_2, text_2 = self.process_file(ruteFile2, '2')
        par_1 = self.parag_div(text_1)
        par_2 = self.parag_div(text_2)
        for p1 in par_1:
            if(p1.strip() != ''):
                sent_1 = self.sent_div(p1)
                for s1 in sent_1:
                    if(s1.strip() != ''):
                        for p2 in par_2:
                            if(p2.strip() != ''):
                                sent_2 = self.sent_div(p2)
                                for s2 in sent_2:
                                    if(s2.strip() != ''):
                                        try:
                                            vectorizer = TfidfVectorizer()
                                            sent_clean_1 = s1.lower().replace('-\n', '')
                                            sent_all_clean_1 = sent_clean_1.lower().replace('\n', '').strip()
                                            sent_clean_2 = s2.lower().replace('-\n', '')
                                            sent_all_clean_2 = sent_clean_2.lower().replace('\n', '').strip()
                                            sent1 = self.remove_stop_words(sent_all_clean_1)
                                            sent2 = self.remove_stop_words(sent_all_clean_2)
                                            X = vectorizer.fit_transform([sent1, sent2])
                                            similarity_matrix = cosine_similarity(X, X)
                                            if(similarity_matrix[0, 1] > 0.75):
                                                self.highlightPDF(doc_1, s1.strip())
                                                self.highlightPDF(doc_2, s2.strip())
                                                print("Подождите!....")
                                        except ValueError:
                                            continue
        doc_1.close()
        doc_2.close()

    def execute(self):
        print("En ejecución...")
        self.labelExecute.setText("Выполняя процесс...")
        self.labelExecute.setStyleSheet("background-color: lightgreen")
        start_time = time()
        if(self.error == False):
            sim_file, porcent = self.search_similarity()
            if(len(sim_file) > 1):
                self.similarity(sim_file[0], sim_file[1])
                self.generate_report(sim_file[0], sim_file[1], porcent)
                self.labelExecute.setText("Выполнение завершено")
                self.labelExecute.setStyleSheet("background-color: lightgreen")
        else:
            self.error = False
        elapsed_time = time() - start_time
        print("Proceso terminado en: %.10f segundos." % elapsed_time)

    def parag_div(self, text):
        return re.split(r'[ \t\r\f\v]*\n[ \t\r\f\v]*\n[ \t\r\f\v]*', text)

    def sent_div(self, text):
        return re.split('(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', text)

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

    def getDataPDF(self, file_name, id_doc):
        ret = []
        try:
            pdf_file = fitz.open(file_name)
            if(id_doc != '0'):
                out = "report\document-" + id_doc + ".pdf"
                page_content = self.getTextPDF(pdf_file)
                pdf_file.save(out)
                pdf_file = fitz.open(out)
                ret.append(pdf_file)
            else:
                page_content = self.getTextPDF(pdf_file).replace('-\n', '')
            ret.append(page_content)
        except Exception:
            self.labelExecute.setText("Ошибка при попытке прочитать документ.")
            self.labelExecute.setStyleSheet("background-color: lightsalmon")
            self.error = True
        return ret

    def getDataDocx(self, file_name, id_doc):
        ret = []
        try:
            doc = docx.Document(file_name)
            if(id_doc != '0'):
                out = "report\document-" + id_doc + ".pdf"
                convert(file_name, out)
                pdf_file = fitz.open(out)
                fullText = self.getTextPDF(pdf_file)
                ret.append(pdf_file)
                ret.append(fullText)
            else:
                fullText = []
                for para in doc.paragraphs:
                    fullText.append(para.text)
                ret.append('\n'.join(fullText))
        except Exception:
            self.labelExecute.setText("Ошибка при попытке прочитать документ.")
            self.labelExecute.setStyleSheet("background-color: lightsalmon")
            self.error = True
        return ret

    def clean_text(self, text):
        text = text.replace('\n', '')
        text = text.replace('-\n', '')
        text = text.replace('-\r', '')
        return text

    def getDataDoc(self, file_name, id_doc):
        ret = []
        try:
            wdFormatPDF = 17
            out = "report\document-" + id_doc + ".pdf"
            out_file = os.path.abspath(out)
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(file_name, False, False, False)
            if(id_doc != '0'):
                doc.SaveAs2(out_file, FileFormat=wdFormatPDF)
                doc.Close()
                word.Quit()
                pdf = fitz.open(out)
                ret.append(pdf)
                fullText = self.getTextPDF(pdf)
            else:
                fullText = doc.Range().Text
                doc.Close()
                word.Quit()
            ret.append(fullText)
        except Exception:
            self.labelExecute.setText("Ошибка при попытке прочитать документ.")
            self.labelExecute.setStyleSheet("background-color: lightsalmon")
            self.error = True
        return ret

    def process_file(self, file_name, id_doc):
        doc, ext = os.path.splitext(file_name)
        if ext == ".doc":
            file_content = self.getDataDoc(file_name, id_doc)
        elif ext == ".docx":
            file_content = self.getDataDocx(file_name, id_doc)
        else:
            file_content = self.getDataPDF(file_name, id_doc)
        return file_content

    def search_max(self, matrix):
        ret = [0, 0, 0]
        max_elem = 0
        for i in range(len(matrix)):
            for j in range(len(matrix[i])):
                elem = matrix[i][j]
                if(elem > max_elem and i != j):
                    max_elem = elem
                    ret[0] = i
                    ret[1] = j
                    ret[2] = round(matrix[i][j]*100,2)
        return ret
    def add_legend(self, draw_obj, chart, data):
        
        legend = Legend()
        legend.alignment = 'right'
        legend.x = 240
        legend.y = 60
        legend.dx = 6.5
        legend.dy = 6.5
        legend.yGap = 0
        legend.deltax = 10
        legend.deltay = 10
        legend.fontName = 'DejaVuSerif'  
        styles = getSampleStyleSheet()
    
        styleBH = styles["Normal"] 
        styleBH.alignment = TA_CENTER
        legend.colorNamePairs = Auto(chart=chart)
        
        draw_obj.add(legend)

    def pie_chart_with_legend(self, data):
        data = [data, 100-data]
        drawing = Drawing(width=400, height=100)
        pie = Pie()
        pie.x = 105
        pie.y = 0
        pie.data = data
        pie.slices.fontColor = None
        pie.labels = 'Сходство Разница'.split()
        pie.slices.strokeWidth = 1
        pie.slices.strokeWidth = 0.5   
        pie.slices[0].strokeColor= colors.HexColor("#d32f2f")
        pie.slices[1].strokeColor= colors.HexColor("#ef9a9a")
        pie.slices[0].fillColor = colors.HexColor("#d32f2f")
        pie.slices[1].fillColor = colors.HexColor("#ef9a9a")
        drawing.add(pie)
        self.add_legend(drawing, pie, data)
        return drawing
    
    def generate_report(self, rute1, rute2, porcent):
        pdfmetrics.registerFont(TTFont('DejaVuSerif','DejaVuSerif.ttf', 'UTF-8'))
        doc = SimpleDocTemplate("report/report.pdf", pagesize=A4, rightMargin=30,leftMargin=30, topMargin=30,bottomMargin=18)
        doc.pagesize = landscape(A4)
        elements = []
        #Graf
        graf = self.pie_chart_with_legend(porcent)
        #Data
        styles = getSampleStyleSheet()
        
        styleBH = styles["Normal"]
        styleBH.alignment = TA_CENTER
        styleBH.fontSize  = 12

        p_head = Paragraph('<font name="DejaVuSerif">Отчет о сходстве между документами </font>', styleBH) 
        p_doc1 = Paragraph('<font name="DejaVuSerif">Документ № 1</font>', styleBH)
        p_doc2 = Paragraph('<font name="DejaVuSerif">Документ № 2</font>', styleBH)

        stylePORC = styles["Normal"]
        stylePORC.alignment = TA_CENTER
        stylePORC.fontSize  = 16
        stylePORC.textColor = colors.HexColor("#d32f2f")

        str_porc = str(porcent) + "%" + " Сходство между документами"
        p_porc = Paragraph('<font name="DejaVuSerif">'+str_porc+'</font>', stylePORC)
        data= [
        [p_head, ""],
        [p_doc1, p_doc2],
        [Paragraph(str(rute1)), Paragraph(str(rute2))], 
        [graf, p_porc]]
        #Tabla
        style = TableStyle([
                       ('SPAN',(0,0),(1,0)),
                       ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                       ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                       ('ALIGN', (0,-1),(-1,-1), 'CENTER'),
                       ('VALIGN', (0,-1),(-1,-1), 'MIDDLE'),
                       ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#29b6f6")),
                       ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#81d4fa"))
                       ])

        #Configure style and word wrap
        s = getSampleStyleSheet()
        s = s["BodyText"]
        s.wordWrap = 'CJK'
        t=Table(data,colWidths=270)
        t.setStyle(style)

        #Send the data and build the file
        elements.append(t)
        doc.build(elements)

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    ventana = VentanaPrincipal()
    ventana.show()
    app.exec_()