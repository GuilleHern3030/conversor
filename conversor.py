from docx2pdf import convert as convertWordToPDF # pip install reportlab docx2pdf
from PIL import Image as PIL_Image # pip install pillow
import pandas as pd # pip install pandas

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import threading

filePath = []
filesLabel = []
converting = False

def convertImageToWebp(fileIn, fileOut):
    #imagen = PIL_Image.open(fileIn)
    PIL_Image.open(fileIn).save(fileOut, 'webp', lossless=True)

def convert(fileIn, fileOut):
    print(f"Converting {fileIn} into {fileOut}")
    if isExtension(fileIn, "png", "jpg") and isExtension(fileOut, "webp"): PIL_Image.open(fileIn).save(fileOut, 'webp', lossless=True)
    elif isExtension(fileIn, "doc") and isExtension(fileOut, "pdf"): convertWordToPDF(fileIn, fileOut)
    else:
        df = None
        if isExtension(fileIn, "xls"): df = pd.read_excel(fileIn)
        elif isExtension(fileIn, "json"): df = pd.read_json(fileIn)
        elif isExtension(fileIn, "csv"): df = pd.read_csv(fileIn)
        else: invalidExtension
        if isExtension(fileOut, "xls"): df.to_excel(fileOut, index=False)
        elif isExtension(fileOut, "json"): df.to_json(fileOut, orient='records')
        elif isExtension(fileOut, "csv"): createCSV(df, fileOut)
        else: invalidExtension

def createCSV(df, fileOut):
    import csv  
    def format_cell(cellContent:str):
        cell = f'{cellContent}'
        if cellContent == None or cell == '' or cell.lower() == 'null' or cell.lower() == 'nan': return "NULL"
        else: return f'"{cell}"'
    df = df.map(format_cell)
    column_names = df.columns
    colmn_names_quoted = ['"' + col + '"' for col in column_names]
    df.columns = colmn_names_quoted
    df.to_csv(fileOut, index=False, quoting=csv.QUOTE_NONE, sep=',')


# Save file
def saveFile(inputFile:str, outputFilePath:str) -> bool:
    try:
        convert(inputFile, outputFilePath)
        labelConvertResult(True)
    except Exception as err: 
        labelConvertResult(False)
        print(err)

def saveFiles(files:list, folderPath:str, extension:str) -> bool:
    try:
        for file in files:
            outputFile = f"{folderPath}/{fileName(file)}.{extension}"
            convert(file, outputFile)
        labelConvertResult(True)
    except Exception as err: 
        labelConvertResult(False)
        print(err)


# Files name
def fileName(aFilePath:str): return (aFilePath.split("/")[-1]).split(".")[0]
def fileExtension(aFilePath:str): return aFilePath.split(".")[-1]
def invalidExtension(): raise Exception("Invalid extension")
def setExtension(aFilePath:str, extension:str):
    extensionStart = aFilePath.find(".")
    #validExtension = validFilesToConvert[validFilesToConvert.index(extension)]
    if (extensionStart >= 0):  return f"{aFilePath.split('.')[0]}.{extension.lower()}"
    else: return f"{aFilePath}.{extension.lower()}"
def isExtension(aFilePath:str, *extension):
    _fileExtension = fileExtension(aFilePath)
    for ext in extension:
        if _fileExtension.find(ext) >= 0:
            return True
    return False


# Posible files to convert
def validFilesConverted(fileExtension:bool):
    if isExtension(fileExtension, "png", "jpg"): return (("WEBP", "*.webp"),("WEBP", "*.webp"))
    elif isExtension(fileExtension, "docx"): return (("PDF", "*.pdf"),("PDF", "*.pdf"))
    else: return (("JSON", "*.json"),("CSV", "*.csv"),("Excel", "*.xlsx"))

validFilesToConvert = (
            ("PNG Image", "*.png"),
            ("JPG Image", "*.jpg"),
            ("Word", "*.docx"),
            ("Excel", "*.xlsx"),
            ("CSV", "*.csv"),
            ("JSON", "*.json")
        )


# Frame
root = Tk()
root.title("Multiple Conversor")

frame = Frame(root)
frame.pack(fill="both", expand="True")

headerFrame = Frame(frame)
headerFrame.grid(row=0, column=0, sticky="nsew", padx=10, pady=7)

convertFrame = Frame(frame)
convertFrame.grid(row=1, column=0, sticky="nsew", padx=10, pady=7)

radiobuttonsFrame = Frame(frame)
radiobuttonsFrame.grid(row=2, column=0, sticky="nsew", padx=10, pady=7)
radiobuttonSelected = IntVar()
extensionSelected = StringVar()
radiobuttons = []

documentsFrame = Frame(frame)
documentsFrame.grid(row=3, column=0, sticky="nsew", padx=10, pady=7)

Label(headerFrame, text="Multiple-Converter", font=("Comic Sans MS", 20)).grid(row=0, column=0, sticky="nsew", padx=10, pady=7)
labelConvertFilesResult = Label(convertFrame, font=("Arial", 18), text="File/s converted successfully")
labelConvertingFiles = Label(convertFrame, font=("Arial", 16), text="Converting...")

def labelConvertResult(success:bool):
    global converting
    converting = False
    if success: labelConvertFilesResult.config(text="Process finished", fg="darkgreen")
    else: labelConvertFilesResult.config(text="An error has been ocurred", fg="red")
    labelConvertFilesResult.grid(row=1, column=0, sticky="nsew", padx=10, pady=7, columnspan=2)
    labelConvertingFiles.grid_forget()

def convertFiles():
    global converting
    if not converting:
        if len(filePath) > 1:
            extensionToSave = validFilesConverted(fileExtension(filePath[0]))[radiobuttonSelected.get()][1].lstrip("*.")
            folderToSave = filedialog.askdirectory(
                title="Save into...", 
                initialdir="./"
            )
            if len(folderToSave) > 0: 
                labelConvertingFiles.grid(row=0, column=1, sticky="nsew", padx=10, pady=7)
                labelConvertFilesResult.grid_forget()
                converting = True
                thread = threading.Thread(target=saveFiles, args=(filePath, folderToSave, extensionToSave))
                thread.start()
        elif len(filePath) == 1:
            typeFiles = validFilesConverted(fileExtension(filePath[0]))
            outputFile = filedialog.asksaveasfilename(
                title="Save as...",
                initialdir=f"./{filePath[0]}",
                initialfile=f"{fileName(filePath[0])}",
                confirmoverwrite=True,
                filetypes=typeFiles,
                typevariable=extensionSelected,
            )
            if len(outputFile) > 0: 
                outputFile = setExtension(outputFile, extensionSelected.get())
                labelConvertingFiles.grid(row=1, column=0, sticky="nsew", padx=10, pady=7)
                labelConvertFilesResult.grid_forget()
                converting = True
                thread = threading.Thread(target=saveFile, args=(filePath[0], outputFile))
                thread.start()
        else: messagebox.showwarning("", "No file selected")
bConvertFiles = Button(convertFrame, text="Convert", font=("Arial", 16), command=convertFiles)
bConvertFiles.config(cursor="hand2", bg="lightgreen")

def selectFiles():
    if not converting:
        results = filedialog.askopenfilenames(
            title="Select", 
            initialdir="./", 
            filetypes=validFilesToConvert
        )
        global radiobuttons
        if len(results) > 0:
            labelConvertFilesResult.grid_forget()
            labelConvertingFiles.grid_forget()
            for radiobutton in radiobuttons: radiobutton.grid_forget()
            radiobuttons = []
            filePath.clear()
            index = 0
            for label in filesLabel: label.grid_forget()
            for result in list(results): 
                filePath.append(result)
                filesLabel.append(Label(documentsFrame, text=result, font=("Arial", 12)))
                filesLabel[-1].grid(row=index, column=0, sticky="we", padx=10, pady=7)
                index += 1
            bConvertFiles.grid(row=0, column=0, sticky="we", padx=10, pady=7)
        if len(results) > 1:
            value = 0
            validExtensions = validFilesConverted(fileExtension(results[0]))
            for extension in validExtensions:
                rb = Radiobutton(
                    radiobuttonsFrame, 
                    text=extension[0], 
                    variable=radiobuttonSelected, 
                    value=value
                )
                rb.grid(row=0, column=value, sticky="nsew", padx=5, pady=7)
                radiobuttons.append(rb)
                value += 1
            radiobuttons[0].select()
bSelectFiles = Button(headerFrame, text="Select", font=("Arial", 16), command=selectFiles)
bSelectFiles.config(cursor="hand2", bg="lightblue")
bSelectFiles.grid(row=1, column=0, sticky="we", padx=10, pady=7)

# Configurando la fila y columna que se adaptarán al tamaño del frame
frame.grid_columnconfigure(0, weight=1)
headerFrame.grid_columnconfigure(0, weight=1)
convertFrame.grid_columnconfigure(0, weight=1)
#radiobuttonsFrame.grid_columnconfigure(0, weight=1)

# Estableciendo el tamaño mínimo de la interfaz
root.update()
root.minsize(frame.winfo_width(), 300) # winfo obtiene tamaño actual
root.mainloop()