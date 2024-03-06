from PyQt5.QtWidgets import *
from PyQt5 import uic, QtWidgets, QtCore
import pandas as pd
import os.path
import os
import win32com.client
import datetime

import pikepdf
from pikepdf import Pdf
import PyPDF2
import xlsxwriter
import mupdf
from pathlib import Path
import openpyxl

import shutil
import fitz

import math
import decimal
from PyPDF2 import PdfFileMerger, PdfFileReader
from decimal import Decimal as D
context = decimal.getcontext()
context.rounding = decimal.ROUND_HALF_UP

#Para arreglar textos
import ftfy
from tqdm import tqdm

from Gui import Ui_MainWindow
Ui_MainWindow, QtBaseClass = uic.loadUiType('Gui.ui')

import qtmodern.styles
import qtmodern.windows

###################################### Recursos #################################################
# https://medium.com/towards-data-science/how-to-easily-convert-a-python-script-to-an-executable-file-exe-4966e253c7e9
# https://stackoverflow.com/questions/40813395/pyqt5-typeerror-wrong-base-class-of-toplevel-widget
# https://www.youtube.com/watch?v=865Q41omqPk
# https://www.codeforests.com/2021/05/16/python-reading-email-from-outlook-2/
# Encapsular https://www.youtube.com/watch?v=3CKYvLW5U7I
# https://medium.com/@akhileshjoshi123/merge-pdfs-with-python-d4d3bfbdbd3b
# En algún momento evaluar utilizarel statement continune despues de los else: (https://www.youtube.com/watch?v=2JsGiygzi5M)
# pyinstaller --onefile --name "APP BOLETAS (NUEVO)" --hiddenimport win32timezone -F main.py
# pyinstaller --onefile --name "APP BOLETAS (NUEVO)" --hiddenimport win32timezone -F --add-data "Gui.ui;ui" main.py

class MyGUI(QMainWindow):

    def __init__(self):
        super(MyGUI, self).__init__()
        uic.loadUi(r'Gui.ui', self)
        title = "Agencia Nacional de Investigación y Desarrollo"
        # set the title
        self.setWindowTitle(title)
        self.show()
        # Aqui van los botones
        self.pushButton.clicked.connect(self.buscar)
        self.pushButton_2.clicked.connect(self.PDF)
        self.pushButton_3.clicked.connect(self.copy_and_rename)
        self.pushButton_4.clicked.connect(self.fusionar)
        self.pushButton_5.clicked.connect(self.merge_excel)
        self.pushButton_6.clicked.connect(self.seleccionar)
        self.dateEdit_2.setDate(datetime.datetime.now().date())

    def show_line_mail(self):
        print(self.lineEdit.text())
        mail = str(self.lineEdit.text())
        return mail

    def show_line_inbox(self):
        print(self.lineEdit_2.text())
        inbox = str(self.lineEdit_2.text())
        return inbox

    def show_line_sender(self):
        print(self.lineEdit_3.text())
        sender = str(self.lineEdit_3.text())
        return sender

    def show_line_subject(self):
        print(self.lineEdit_4.text())
        subject = str(self.lineEdit_4.text())
        return subject

    def show_spin(self):
        print(self.spinBox.text())
        number = int(self.spinBox.text())
        return number

    def date_start(self):
        print(self.dateEdit.text())
        self.dateEdit.setDisplayFormat("yyyy-dd-MM")
        start_time = self.dateEdit.text()
        print(start_time)
        return start_time

    def date_end(self):
        print(self.dateEdit_2.text())
        self.dateEdit_2.setDisplayFormat("yyyy-dd-MM")
        end_time = self.dateEdit_2.text()
        print(end_time)
        return end_time

    def buscar(self):
        output_dir = Path.cwd() / "Boletas (PDF)"
        output_dir.mkdir(parents=True, exist_ok=True)

        ########################################### Filtrar por Fecha #######################################################
        '''IMPORTANTE el filtro [ReceivedTime] solicita AÑO-DIA-MES, caso contrario toma cualquier rango de fechas'''

        mail = self.show_line_mail()
        inbox = self.show_line_inbox()
        sender = self.show_line_sender()
        subject = self.show_line_subject()

        start_time = self.date_start()  # AÑO-DIA-MES
        end_time = self.date_end()

        if mail == "" and inbox == "" and sender == "" and subject == "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            #inbox = outlook.Folders(self.show_line_mail).Folders(self.show_line_inbox)
            messages = inbox.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
        elif mail != "" and inbox == "" and sender == "" and subject == "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(mail)).Folders('Bandeja de entrada')
            # .Folders(self.show_line_inbox)
            messages = inbox.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
        elif mail == "" and inbox != "" and sender == "" and subject == "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(outlook.Folders.Item(1))).Folders(str(inbox))
            messages = inbox.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
        elif mail == "" and inbox == "" and sender != "" and subject == "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            #inbox = outlook.Folders(self.show_line_mail).Folders(self.show_line_inbox)
            messages = inbox.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "' And [SenderEmailAddress] = '" + sender + "'")
        elif mail == "" and inbox == "" and sender == "" and subject != "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            #inbox = outlook.Folders(self.show_line_mail).Folders(self.show_line_inbox)
            messages = inbox.Items
            messages = messages.Restrict(f"[Subject] = '{subject}'" + " And [ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
        elif mail != "" and inbox != "" and sender == "" and subject == "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(mail)).Folders(str(inbox))
            messages = inbox.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
        elif mail != "" and inbox == "" and sender != "" and subject == "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(mail)).Folders('Bandeja de entrada')
            messages = inbox.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "' And [SenderEmailAddress] = '" + sender + "'")
        elif mail != "" and inbox == "" and sender == "" and subject != "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(mail)).Folders('Bandeja de entrada')
            messages = inbox.Items
            messages = messages.Restrict(f"[Subject] = '{subject}'" + "And [ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
        elif mail == "" and inbox != "" and sender != "" and subject == "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(inbox))
            messages = inbox.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "' And [SenderEmailAddress] = '" + sender + "'")
        elif mail == "" and inbox != "" and sender == "" and subject != "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(inbox))
            messages = inbox.Items
            messages = messages.Restrict(f"[Subject] = '{subject}'" + "' And [ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
        elif mail == "" and inbox == "" and sender != "" and subject != "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            # inbox = outlook.Folders(self.show_line_mail).Folders(self.show_line_inbox)
            messages = inbox.Items
            messages = messages.Restrict(f"[Subject] = '{subject}'" + "And [ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "' And [SenderEmailAddress] = '" + sender + "'")
        elif mail != "" and inbox != "" and sender != "" and subject == "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(mail)).Folders(str(inbox))
            messages = inbox.Items
            messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "' And [SenderEmailAddress] = '" + sender + "'")
        elif mail != "" and inbox != "" and sender == "" and subject != "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(mail)).Folders(str(inbox))
            messages = inbox.Items
            messages = messages.Restrict(f"[Subject] = '{subject}'" + "' And [ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
        elif mail != "" and inbox == "" and sender != "" and subject != "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(mail)).Folders('Bandeja de entrada')
            messages = inbox.Items
            messages = messages.Restrict(f"[Subject] = '{subject}'" + "' And [ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "' And [SenderEmailAddress] = '" + sender + "'")
        elif mail == "" and inbox != "" and sender != "" and subject != "":
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.Folders(str(inbox))
            messages = inbox.Items
            messages = messages.Restrict(f"[Subject] = '{subject}'" + "' And [ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "' And [SenderEmailAddress] = '" + sender + "'")

        ########################################################################################################################
        # filtro = "[SenderEmailAddress] = '0m3r@email.com'"

        # Items = inbox.Items.Restrict(filter)
        # Item = Items.GetFirst()
        ########################################################################################################################
        print("messages")
        #(activar i para enumerar pdf descargados  # attachment_name = "{:04n})
        Nº = []
        correos = []
        fechas = []
        adjunto = []
        mensaje = []
        i = 1
        for message in messages:
            if message.Class == 43:
                current_sender = str(message.Sender).lower()
                mail = str(message.SenderEmailAddress).lower()
                subject = message.Subject
                body = message.body
                attachments = message.Attachments
                attachments_name = str(attachments).lower()
                #Path(target_folder / "EMAIL_BODY_{}.txt").write_text(str(body))
                if message.SenderEmailType == "EX":
                    print(message.Sender.GetExchangeUser().PrimarySmtpAddress)
                    print(message.Class)
                    print(message.SenderEmailType)
                    print(message.ReceivedTime)
                    mail = str(message.Sender.GetExchangeUser().PrimarySmtpAddress).lower()
                    date = str(message.ReceivedTime)[0:19]
                    date = datetime.datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
                    date = date.strftime('%d-%m-%Y %H:%M:%S')
                    for attachment in attachments:
                        if ".pdf" in str(attachment):
                        #if ".pdf" in str(attachment) and "bhe" in str(attachment):
                            target_folder = output_dir
                            target_folder.mkdir(parents=True, exist_ok=True)
                            print(message.Class)
                            print(target_folder)
                            print(attachment)
                            print(mail)
                            attachment_name = str(attachment).lower()
                            attachment_name = "{:04n} - {attachment_name}".format(i, attachment_name=attachment_name)
                            attachment.SaveAsFile(target_folder / attachment_name)
                            Nº.append(i)
                            correos.append(mail)
                            fechas.append(date)
                            adjunto.append(attachment_name)
                            mensaje.append(str(body))
                            i += 1
                else:
                    print(message.SenderEmailAddress)
                    print(message.Class)
                    print(message.SenderEmailType)
                    print(message.ReceivedTime)
                    mail = str(message.SenderEmailAddress).lower()
                    date = str(message.ReceivedTime)[0:19]
                    date = datetime.datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
                    date = date.strftime('%d-%m-%Y %H:%M:%S')
                    for attachment in attachments:
                        if ".pdf" in str(attachment):
                        #if ".pdf" in str(attachment) and "bhe" in str(attachment):
                            target_folder = output_dir
                            target_folder.mkdir(parents=True, exist_ok=True)
                            print(message.Class)
                            print(target_folder)
                            print(attachment)
                            print(mail)
                            attachment_name = str(attachment).lower()
                            attachment_name = "{:04n} - {attachment_name}".format(i, attachment_name=attachment_name)
                            attachment.SaveAsFile(target_folder / attachment_name)
                            Nº.append(i)
                            correos.append(mail)
                            fechas.append(date)
                            adjunto.append(attachment_name)
                            mensaje.append(str(body))
                            i += 1
        self.dateEdit.setDisplayFormat("dd-MM-yyyy")
        self.dateEdit_2.setDisplayFormat("dd-MM-yyyy")
        df = pd.DataFrame()
        df["Nº"] = Nº
        df["Correo"] = correos
        df["Fecha de envío"] = fechas
        df["PDF"] = adjunto
        df["Mensaje"] = mensaje
        writer = pd.ExcelWriter('DETALLE ENVIOS.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Boletas de Honorarios (Detalle)', encoding='ascii', index=False)
        writer.save()

        dlg = QMessageBox(self)
        dlg.setWindowTitle("Outlook")
        dlg.setText("La busqueda de Boletas de Honorarios ha finalizado.")
        button = dlg.exec_()
        if button == QMessageBox.Ok:
            print("OK!")

    def PDF(self):
        direccion = Path.cwd() / "Boletas (PDF)"
        direccion.mkdir(parents=True, exist_ok=True)

        def listarArchivos(direccion):
            listaPDF = []
            nombreArchivos = os.listdir(direccion)
            for archivo in nombreArchivos:
                if ".pdf" in archivo:
                    listaPDF.append(archivo)
            # listaPdf.sort(key=lambda x: os.path.getmtime(os.path.join(direccion, x)))
            return listaPDF

        listaPDF = listarArchivos(direccion)
        df = pd.DataFrame()
        df["N"] = [i + 1 for i in range(len(listaPDF))]
        df["PDF"] = pd.DataFrame(listaPDF)
        df["TEXTO"] = ""
        df["Boleta"] = ""
        df["Rut Emisor"] = ""
        df["Nº Boleta"] = ""
        df["Nombre ANID"] = ""
        df["RUT ANID"] = ""
        df["Direccion ANID"] = ""
        df["Total Honorarios"] = ""
        df["Impuestos"] = ""
        df["Total"] = ""
        df["% Impuesto Retenido"] = ""
        df["Fecha de Boleta"] = ""
        df["Fecha de Emisión"] = ""
        df["Detalle"] = ""

        for index, row in tqdm(df.iterrows(), total=len(df)):
            print(str("número ") + str(int(index) + 1) + str(" de ") + str(len(df["PDF"])) + " | Boleta Nombre: " + str(df.at[index, "PDF"]))
            target_folder = direccion / df.at[index, "PDF"]
            filename = target_folder
            pdfFile = open(filename, 'rb')  # open function reads the file
            pdfReader = PyPDF2.PdfFileReader(pdfFile, strict=False)
            # The pdfReader variable is a readable object that will be
            parsedpageCt = pdfReader.numPages
            count = 0
            text = ""
            # The while loop will read each page
            while count < parsedpageCt:
                print(parsedpageCt)
                pageObj = pdfReader.getPage(count)
                count = count + 1
                text += pageObj.extractText()
            text = ftfy.fix_text(text)
            df.at[index, 'TEXTO'] = text
            print(text)
            ####################  NOMBRE  #########################
            if 'BOLETA ' in text:
                #    start = text.index('$')
                end = text.index('BOLETA', 0 + 1)
                NOMBRE = text[0:end]
                print(f"Start: {0}, End: {end}")
                print(NOMBRE)
                df.at[index, "Boleta"] = NOMBRE
            else:
                NOMBRE = "-"
                print(NOMBRE)
                df.at[index, "Boleta"] = NOMBRE
            ################# NUMERO DE BOLETA ###################
            if 'N ° ' in text and 'RUT:' in text:
                start = text.index('N ° ') + len('N ° ')
                end = text.index('RUT:', start + 1)
                NUMERO_BOLETA = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(NUMERO_BOLETA)
                df.at[index, "Nº Boleta"] = NUMERO_BOLETA
            else:
                NUMERO_BOLETA = "-"
                print(NUMERO_BOLETA)
                df.at[index, "Nº Boleta"] = NUMERO_BOLETA
            ################## Fecha INGRESADA ###################
            if 'Fecha:' in text and 'Señor(es):' in text:
                start = text.index('Fecha:') + len('Fecha:')
                end = text.index('Señor(es):', start + 1)
                FECHA_EVALUADOR = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(FECHA_EVALUADOR)
                df.at[index, "Fecha de Boleta"] = FECHA_EVALUADOR
            else:
                FECHA_EVALUADOR = "-"
                print(FECHA_EVALUADOR)
                df.at[index, "Fecha de Boleta"] = FECHA_EVALUADOR
            ################## Nombre ANID ###################
            if 'Señor(es):' in text and 'Rut:' in text:
                start = text.index('Señor(es):') + len("Señor(es):")
                end = text.index('Rut:', start + 1)
                NOMBRE_ANID = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(NOMBRE_ANID)
                df.at[index, "Nombre ANID"] = NOMBRE_ANID
            else:
                NOMBRE_ANID = "-"
                print(NOMBRE_ANID)
                df.at[index, "Nombre ANID"] = NOMBRE_ANID
            ################## Rut ANID ###################
            if 'Rut:' in text and 'Domicilio:' in text:
                start = text.index('Rut:') + len("Rut:")
                end = text.index('Domicilio:', start + 1)
                RUT_ANID = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(RUT_ANID)
                df.at[index, "RUT ANID"] = RUT_ANID
            else:
                RUT_ANID = "-"
                print(RUT_ANID)
                df.at[index, "RUT ANID"] = RUT_ANID
            ################## DIRECCION ANID ###################
            if 'Domicilio:' in text and 'Por atención profesional:' in text:
                start = text.index('Domicilio:') + len('Domicilio:')
                end = text.index('Por atención profesional:', start + 1)
                DIRECCION_ANID = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(DIRECCION_ANID)
                df.at[index, "Direccion ANID"] = DIRECCION_ANID
            else:
                DIRECCION_ANID = "-"
                print(DIRECCION_ANID)
                df.at[index, "Direccion ANID"] = DIRECCION_ANID
            ################## RUT EMISOR #######################
            if 'RUT:' in text and 'GIRO(S)' in text:
                start = text.index('RUT:') + len('RUT:')
                end = text.index('GIRO(S)', start + 1)
                RUT_EMISOR = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(RUT_EMISOR)
                df.at[index, "Rut Emisor"] = RUT_EMISOR
            else:
                RUT_EMISOR = "-"
                print(RUT_EMISOR)
                df.at[index, "Rut Emisor"] = RUT_EMISOR
            ################## HONORARIOS ###################
            if 'Total Honorarios $:' in text and ('10.75 % Impto.' or '11.50 % Impto.' or '11.5 % Impto.' or
                                                  '12.25 % Impto.' or '13 % Impto.' or '13.0 % Impto.' or '13.00 % Impto.' or '13.75 % Impto.' or
                                                  '14.5 % Impto.' or '14.50 % Impto.' or '15.25 % Impto.' or '16 % Impto.' or '16.0 % Impto.' or
                                                  '16.00 % Impto.' or '16.75 % Impto.' or'17 % Impto.' or '17.0 % Impto.' or '17.00 % Impto.' or
                                                  '17.75 % Impto.' or '18.25 % Impto.' or '19 % Impto.' or '19.0 % Impto.' or '19.00 % Impto.' or
                                                  '20 % Impto.' or '20.0 % Impto.' or '20.00 % Impto.' in text):
                start = text.index('Total Honorarios $:') + len("Total Honorarios $:")
                if '10.75 % Impto.' in text:
                    end = text.index('10.75 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '10.75 %'
                    print(f"Start: {start}, End: {end}")
                    print('10.75 %')
                elif '11.5 % Impto.' in text:
                    end = text.index('11.5 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '11.5 %'
                    print(f"Start: {start}, End: {end}")
                    print('11.5 %')
                elif '11.50 % Impto.' in text:
                    end = text.index('11.50 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '11.50 %'
                    print(f"Start: {start}, End: {end}")
                    print('11.50 %')
                elif '12.25 % Impto.' in text:
                    end = text.index('12.25 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '12.25 %'
                    print(f"Start: {start}, End: {end}")
                    print('12.25 %')
                elif '13 % Impto.' in text:
                    end = text.index('13 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '13 %'
                    print(f"Start: {start}, End: {end}")
                    print('13 %')
                elif '13.0 % Impto.' in text:
                    end = text.index('13.0 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '13.0 %'
                    print(f"Start: {start}, End: {end}")
                    print('13.0 %')
                elif '13.00 % Impto.' in text:
                    end = text.index('13.00 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '13.00 %'
                    print(f"Start: {start}, End: {end}")
                    print('13.00 %')
                elif '13.75 % Impto.' in text:
                    end = text.index('13.75 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '13.75 %'
                    print(f"Start: {start}, End: {end}")
                    print('13.75 %')
                elif '14.5 % Impto.' in text:
                    end = text.index('14.5 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '14.5 %'
                    print(f"Start: {start}, End: {end}")
                    print('14.5 %')
                elif '14.50 % Impto.' in text:
                    end = text.index('14.50 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '14.50 %'
                    print(f"Start: {start}, End: {end}")
                    print('14.50 %')
                elif '15.25 % Impto.' in text:
                    end = text.index('15.25 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '15.25 %'
                    print(f"Start: {start}, End: {end}")
                    print('15.25 %')
                elif '16 % Impto.' in text:
                    end = text.index('16 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '16 %'
                    print(f"Start: {start}, End: {end}")
                    print('16 % Impto.')
                elif '16.0 % Impto.' in text:
                    end = text.index('16.0 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '16.0 %'
                    print(f"Start: {start}, End: {end}")
                    print('16.0 % Impto.')
                elif '16.00 % Impto.' in text:
                    end = text.index('16.00 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '16.00 %'
                    print(f"Start: {start}, End: {end}")
                    print('16.00 % Impto.')
                elif '16.75 % Impto.' in text:
                    end = text.index('16.75 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '16.75 %'
                    print(f"Start: {start}, End: {end}")
                    print('16.75 % Impto.')
                elif '17 % Impto.' in text:
                    end = text.index('17 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '17 %'
                    print(f"Start: {start}, End: {end}")
                    print('17 %')
                elif '17.0 % Impto.' in text:
                    end = text.index('17.0 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '17.0 %'
                    print(f"Start: {start}, End: {end}")
                    print('17.0 %')
                elif '17.00 % Impto.' in text:
                    end = text.index('17.00 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '17.00 %'
                    print(f"Start: {start}, End: {end}")
                    print('17.00 %')
                elif '17.75 % Impto.' in text:
                    end = text.index('17.75 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '17.75 %'
                    print(f"Start: {start}, End: {end}")
                    print('17.75 %')
                elif '18.25 % Impto.' in text:
                    end = text.index('18.25 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '18.25 %'
                    print(f"Start: {start}, End: {end}")
                    print('18.25 %')
                elif '19 % Impto.' in text:
                    end = text.index('19 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '19 %'
                    print(f"Start: {start}, End: {end}")
                    print('19 %')
                elif '19.0 % Impto.' in text:
                    end = text.index('19.0 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '19.0 %'
                    print(f"Start: {start}, End: {end}")
                    print('19.0 %')
                elif '19.00 % Impto.' in text:
                    end = text.index('19.00 % Impto.', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '19.00 %'
                    print(f"Start: {start}, End: {end}")
                    print('19.00 %')
                else:
                    end = text.index('Fecha / Hora', start + 1)
                    df.at[index, "% Impuesto Retenido"] = '-'
                    print(f"Start: {start}, End: {end}")
                    print('-')
                if end > start:  # Significa que encontró el final y es mayor a 'Total Honorarios $:', si no encuentra la retenciòn ocupa el "Fecha / Hora" que es el siguiente texto que viene cuando no hay retención.
                    HONORARIOS = text[start:end]
                    print(f"Start: {start}, End: {end}")
                    print(HONORARIOS)
                    df.at[index, "Total Honorarios"] = HONORARIOS
                else:  # En este else no encontrò ni el '%' ni 'Fecha /hora'.
                    HONORARIOS = "-"
                    print(f"Start: {start}, End: {end}")
                    print(HONORARIOS)
                    df.at[index, "Total Honorarios"] = HONORARIOS
            else:  # Si no encuentra ningun 'Total Honorarios $:' entonces incorpora un "-" a honorarios y a % impuesto retenido.
                df.at[index, "% Impuesto Retenido"] = '-'
                df.at[index, "Total Honorarios"] = '-'
                print(df.at[index, "% Impuesto Retenido"])
                print(df.at[index, "Total Honorarios"])
            ################## IMPUESTO RETENIDO ###################
            if 'Impto. Retenido:' in text and 'Total:' in text:
                start = text.index('Impto. Retenido:') + len('Impto. Retenido:')
                end = text.index('Total:', start + 1)
                IMPUESTO_RETENIDO = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(IMPUESTO_RETENIDO)
                df.at[index, "Impuestos"] = IMPUESTO_RETENIDO
            else:
                IMPUESTO_RETENIDO = "-"
                print(IMPUESTO_RETENIDO)
                df.at[index, "Impuestos"] = IMPUESTO_RETENIDO
            ################## TOTAL ###################
            if 'Total:' in text and ('Esta boleta tiene una retención' or 'Fecha / Hora' in text):
                start = text.index('Total:') + len('Total:')
                if 'Esta boleta tiene una retención' in text:
                    end = text.index('Esta boleta tiene una retención', start + 1)
                else:
                    end = text.index('Fecha / Hora', start + 1)
                TOTAL = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(TOTAL)
                df.at[index, "Total"] = TOTAL
            else:
                TOTAL = "-"
                print(TOTAL)
                df.at[index, "Total"] = TOTAL
            ################## FECHA_EMISION ##################
            if 'Fecha / Hora' in text:
                start = text.index('Fecha / Hora') + len('Fecha / Hora Emisión: ')
                end = text.index(':', start + 1) - 3
                FECHA_EMISION = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(FECHA_EMISION)
                df.at[index, "Fecha de Emisión"] = FECHA_EMISION
            else:
                FECHA_EMISION = "-"
                print(FECHA_EMISION)
                df.at[index, "Fecha de Emisión"] = FECHA_EMISION
            ################## DETALLE ##################
            if 'Por atenci' in text and 'Total Honorarios $:' in text:
                start = text.index('Por atenci') + len('Por atención profesional:')
                end = text.index('Total Honorarios $:', start + 1) - len(TOTAL)
                DETALLE = text[start:end]
                print(f"Start: {start}, End: {end}")
                print(DETALLE)
                df.at[index, "Detalle"] = DETALLE
            else:
                DETALLE = "-"
                print(DETALLE)
                df.at[index, "Detalle"] = DETALLE

        #df = df.replace('\n', '', regex=True)
        df[df.columns[df.columns != 'TEXTO']] = df[df.columns[df.columns != 'TEXTO']].replace('\n', '', regex=True)

        df["Nombre ANID"] = df["Nombre ANID"].str.lstrip()
        df["RUT ANID"] = df["RUT ANID"].str.lstrip()
        df["Direccion ANID"] = df["Direccion ANID"].str.lstrip()
        df["Rut Emisor"] = df["Rut Emisor"].str.lstrip()
        df["Total Honorarios"] = df["Total Honorarios"].str.lstrip()
        df["Impuestos"] = df["Impuestos"].str.lstrip()
        df["Total"] = df["Total"].str.lstrip()
        df["Fecha de Boleta"] = df["Fecha de Boleta"].str.lstrip()

        ################################# COMPROBACION #############################################
        df['Resultado'] = df.apply(lambda x: "REVISAR" if (x["Boleta"] == "-" or
                                                           x["Rut Emisor"] == "-" or
                                                           x["Nº Boleta"] == "-" or
                                                           x["Nombre ANID"] == "-" or
                                                           x["RUT ANID"] == "-" or
                                                           x["Direccion ANID"] == "-" or
                                                           x["Total Honorarios"] == "-" or
                                                           x["Impuestos"] == "-" or
                                                           x["Total"] == "-" or
                                                           x["% Impuesto Retenido"] == "-" or
                                                           x["Fecha de Boleta"] == "-" or
                                                           x["Fecha de Emisión"] == "-" or
                                                           x["Detalle"] == "-") else "LECTURA EXITOSA", axis=1)

        # df.to_csv("BOLETAS.csv", sep=';', encoding='latin-1', index=False, decimal=',')

        dpl = df.groupby(["Boleta", "Rut Emisor", "Nº Boleta"], as_index = False)["PDF"].count()

        writer = pd.ExcelWriter('BOLETAS (PDF).xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Boletas de Honorarios', index=False)
        dpl.to_excel(writer, sheet_name='Conteo por Evaluador', index=False)
        writer.close()

        print("Proceso Finalizado")

        dlg = QMessageBox(self)
        dlg.setWindowTitle("PDF a EXCEL")
        dlg.setText("El archivo ha sido generado con un total de " + str(len(listaPDF)) + " PDFs detectados.")
        button = dlg.exec_()
        if button == QMessageBox.Ok:
            print("OK!")


    def merge_excel(self):
        DETALLE = pd.read_excel("DETALLE ENVIOS.xlsx", thousands='.')
        BOLETAS = pd.read_excel("BOLETAS (PDF).xlsx", thousands='.')
        #BOLETAS[["Total Honorarios", "Impuestos", "Total"]] = BOLETAS[["Total Honorarios", "Impuestos", "Total"]].astype(str)
        df = pd.merge(BOLETAS, DETALLE, on=['PDF'], how='inner')
        df[['Fecha de envío', 'Hora de envío']] = df['Fecha de envío'].str.split(' ', expand=True)
        df = df.fillna('-')
        df = df[["Nº", "Correo", "Fecha de envío", "Hora de envío", "PDF", "Boleta", "Rut Emisor", "Nº Boleta",
                 "Nombre ANID", "RUT ANID", "Direccion ANID", "Total Honorarios", "Impuestos", "Total",
                 "% Impuesto Retenido", "Fecha de Boleta", "Detalle", "Resultado"]]
        writer = pd.ExcelWriter('BOLETAS (TODO).xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Boletas de Honorarios', encoding='ascii', index=False)
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['Boletas de Honorarios'].set_column(col_idx, col_idx, column_width)
        writer.save()

        dlg = QMessageBox(self)
        dlg.setWindowTitle("Fechas de Correos")
        dlg.setText("Se ha incorporado mail, hora y fecha de los adjuntos.")
        button = dlg.exec_()
        if button == QMessageBox.Ok:
            print("OK!")


    def seleccionar(self):
        # Leer el archivo Excel con Pandas
        if not os.path.exists("BOLETAS (SELECCION).xlsx"):
            dlg = QMessageBox(self)
            dlg.setWindowTitle("Selección")
            dlg.setText("Archivo 'BOLETAS (SELECCION).xlsx' no está en la carpeta")
            button = dlg.exec_()
            if button == QMessageBox.Ok:
                print("OK!")
        else:
            df = pd.read_excel("BOLETAS (SELECCION).xlsx")
            if "PDF" in df.columns: # La columna "PDF" existe en el DataFrame
                # Iterar sobre las filas del DataFrame
                no = 0
                for index, row in tqdm(df.iterrows(), total=len(df)):
                    # Obtener el nombre del archivo PDF
                    nombre_archivo = row["PDF"]
                    # Obtener la ruta del archivo PDF
                    ruta_archivo = Path.cwd() / "Boletas (PDF)" / nombre_archivo
                    # Crear la subcarpeta (si no existe)
                    subcarpeta = "Boletas (SELECCION)"
                    if not os.path.exists(subcarpeta):
                        os.mkdir(subcarpeta)
                    # Mover el archivo PDF a la subcarpeta
                    if not os.path.exists(ruta_archivo):
                        print("Boleta no encontrada: ", nombre_archivo)
                        no += 1
                        continue
                    shutil.copy(str(ruta_archivo), Path.cwd() / subcarpeta / str(nombre_archivo))
                dlg = QMessageBox(self)
                dlg.setWindowTitle("Traspaso de boletas")
                dlg.setText(str(len(df) - no) + " boletas traspasadas de carpeta")
                button = dlg.exec_()
                if button == QMessageBox.Ok:
                    print("OK!")
            else:
                # La columna "PDF" no existe en el DataFrame
                dlg = QMessageBox(self)
                dlg.setWindowTitle("Traspaso de boletas")
                dlg.setText("La columna 'PDF' no existe en el archivo 'BOLETAS (SELECCION).xlsx'.")
                button = dlg.exec_()
                if button == QMessageBox.Ok:
                    print("OK!")


    def copy_and_rename(self):
        direccion = Path.cwd() / "Boletas (SELECCION)"
        direccion.mkdir(parents=True, exist_ok=True)

        def listarArchivos(direccion):
            listaPDF = []
            nombreArchivos = os.listdir(direccion)
            for archivo in nombreArchivos:
                if ".pdf" in archivo:
                    listaPDF.append(archivo)
            #listaPDF.sort(key=lambda x: os.path.getmtime(os.path.join(direccion, x)))
            return listaPDF

        listaPDF = listarArchivos(direccion)

        j = 0
        while j < len(listaPDF):
            folder = direccion / listaPDF[j]
            output_dir = Path.cwd() / "Boletas (ENUMERADAS)"
            output_dir.mkdir(parents=True, exist_ok=True)
            filename = folder
            filename_new = output_dir / "{:04n} - BOLETA SCH.pdf".format(j + 1)
            #pdf = Pdf.open(filename)
            #pdf.save(filename_new)
            shutil.copyfile(filename, filename_new)
            print("Boleta " + listaPDF[j] + " renombrada como {:04n} - BOLETA SCH.pdf".format(j+1))
            j += 1
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Enumeración de Archivos")
        dlg.setText(str(len(listaPDF)) + " archivos han sido enumerados.")
        button = dlg.exec_()
        if button == QMessageBox.Ok:
            print("OK!")


    def fusionar(self):
        #https://stackoverflow.com/questions/3444645/merge-pdf-files
        #pip install PyMuPDF
        number = self.show_spin()
        direccion = Path.cwd() / "Boletas (ENUMERADAS)"
        direccion.mkdir(parents=True, exist_ok=True)

        def listarArchivos(direccion):
            listaPDF = []
            nombreArchivos = os.listdir(direccion)
            for archivo in nombreArchivos:
                if ".pdf" in archivo:
                    listaPDF.append(archivo)
            #listaPDF.sort(key=lambda x: os.path.getmtime(os.path.join(direccion, x)))
            return listaPDF

        listaPDF = listarArchivos(direccion)

        pdfs = fitz.Document()
        z = 1
        j = 0
        while j < len(listaPDF):
            file = direccion / listaPDF[j]
            mfile = fitz.Document(file)
            output_dir = Path.cwd() / "Fusion PDFs"
            output_dir.mkdir(parents=True, exist_ok=True)
            x = D(str((j + 1) / number))
            pdfs.insert_pdf(mfile)
            if D(x) == D(z) and D(x) < D(str(math.ceil(len(listaPDF) / number))):
                nombre_archivo_salida = str(output_dir) + "/" + "TED ({}).pdf".format(z)
                print(nombre_archivo_salida)
                pdfs.save(nombre_archivo_salida)
                pdfs = fitz.Document()
                z += 1
                print("Lote " + str(z) + " de " + str(math.ceil(len(listaPDF) / number)))
            elif D(x) == D(str((len(listaPDF) / number))):
                nombre_archivo_salida = output_dir / "TED ({}).pdf".format(math.ceil(len(listaPDF) / number))
                print(nombre_archivo_salida)
                pdfs.save(nombre_archivo_salida)
                print("Lote " + str(math.ceil(len(listaPDF) / number)) + " de " + str(math.ceil(len(listaPDF) / number)))
            j += 1
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Fusión de PDFs")
        dlg.setText("Se han generado " + str(math.ceil(len(listaPDF) / number)) + " fusiones de PDFs para la elaboración de TEDs")
        button = dlg.exec_()
        if button == QMessageBox.Ok:
            print("OK!")


def main():
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QApplication([])
    app.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    window = MyGUI()
    #qtmodern.styles.dark(app)
    qtmodern.styles._apply_base_theme(app)
    mw = qtmodern.windows.ModernWindow(window)
    mw.show()
    app.exec_()

if __name__ == '__main__':
    main()

