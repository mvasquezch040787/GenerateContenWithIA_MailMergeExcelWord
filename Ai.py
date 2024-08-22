import customtkinter.windows
import docx as docx
import google.generativeai as genai
from Data import Texts
import customtkinter 
import tkinter as tk
import openpyxl
from tqdm import tqdm
import time
from tkinter import ttk
import multiprocessing

workBookData=""
sheetTopics=""
sheetData = ""
customtkinter.set_appearance_mode("light")
app = customtkinter.CTk()

app.title("Generate")

lblVacio=customtkinter.CTkLabel(master=app, text="          ")
lblVacio.grid(row=0,column=0)

lblCarrera=customtkinter.CTkButton(master=app, text="Carrera",fg_color="gray")
lblCarrera.grid(row=0,column=1)

TxtCarrera=customtkinter.CTkEntry(master=app, width=1250)
TxtCarrera.grid(row=0,column=2)


lblUnidad=customtkinter.CTkButton(master=app, text="Unidad Didactica ",fg_color="gray")
lblUnidad.grid(row=1,column=1)

TxtUnidad=customtkinter.CTkEntry(master=app, width=1250,placeholder_text="Ingresar la Unidad Didactica")
TxtUnidad.grid(row=1,column=2)

TxtMessage=customtkinter.CTkEntry(master=app, width=1250)
TxtMessage.grid(row=5,column=2)

TxtCarrera.insert(0,"ARQUITECTURA DE PLATAFORMAS Y SERVICIOS DE TECNOLOGÍAS DE INFORMACIÓN")


def seleccionarTemasExcel():
    fileExcelTopics =tk.filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
    workbook = openpyxl.load_workbook(fileExcelTopics)
    global sheetTopics 
    sheetTopics=workbook.active 

def seleccionarArchivoData():
    global fileExcelData
    fileExcelData = tk.filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
    global workBookData
    workBookData=openpyxl.load_workbook(fileExcelData)
    global sheetData
    sheetData=workBookData.active 

def Generate():
    genai.configure(api_key=Texts.API_KEY)
    model=genai.GenerativeModel(model_name="gemini-pro")    
    row_count=2
    for row in sheetTopics.iter_rows(min_row=1,max_row=16,max_col=1):
        for cell in row:
            try:
                if cell.value is None:
                    break
                else:
                    IL=str(model.generate_content(Texts.ACCURACYIL+cell.value+ " De "+TxtUnidad.get()+ "de la carrera de "+TxtCarrera.get()).text)
                    time.sleep(1.5)
                    sheetData[f'A{row_count}']=row_count-1
                    sheetData[f'B{row_count}']=IL
                    sheetData[f'C{row_count}']=row_count-1
                    sheetData[f'D{row_count}']=cell.value
                    sheetData[f'E{row_count}']=(model.generate_content(Texts.ACCURACY_PROPOSITO+cell.value+ " De "+IL).text)
                    time.sleep(1.5)
                    sheetData[f'F{row_count}']=(model.generate_content(Texts.ACCURACYIL_RECURSOS+cell.value+ " De "+IL).text)
                    time.sleep(1.5)
                    sheetData[f'G{row_count}']=(model.generate_content(Texts.ACCURACY_INICIO+cell.value+ " De "+IL).text)
                    time.sleep(1.5)
                    sheetData[f'H{row_count}']=(model.generate_content(Texts.ACCURACY_DESARROLLO+cell.value+ " De "+IL).text)
                    time.sleep(1.5)
                    sheetData[f'I{row_count}']=(model.generate_content(Texts.ACCURACY_CIERRE+cell.value+ " De "+IL).text)
                    time.sleep(1.5)
                    workBookData.save(fileExcelData)
                    row_count+=1
            except ValueError as e:
                print(e)
        if row_count>17:
            TxtMessage.insert(tk.END, f"FELICITACIONES GENERACIÒN SATISFACTORIA")
            break
  
btnTopics=customtkinter.CTkButton(master=app, text="Seleccionar Temas EXCEL", width=1250,command=seleccionarTemasExcel)
btnTopics.grid(row=2,column=2)

BtnChangeXlsx=customtkinter.CTkButton(master=app, text="Seleccionar Almacenamiento EXCEL",width=1250,command=seleccionarArchivoData)
BtnChangeXlsx.grid(row=3,column=2)

BtnGenerate=customtkinter.CTkButton(master=app, text="GENERAR CONTENIDO USANDO IA",width=1250,command=Generate)
BtnGenerate.grid(row=4,column=2)
app.state("zoomed")
app.mainloop()




