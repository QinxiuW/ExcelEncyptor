import os
import glob
import win32com.client as win32
import PySimpleGUI as sg

# Función para encriptar los archivos Excel en un directorio
def encriptar_excel(directorio, password):
    # Obtener la lista de archivos Excel en el directorio
    archivos_excel = glob.glob(os.path.join(directorio, '*.xlsx'))
    print(archivos_excel)
    # Crear una instancia de Excel
    excel = win32.Dispatch('Excel.Application')

    # Iterar sobre cada archivo Excel y protegerlo con contraseña
    for archivo in archivos_excel:
        print(archivo)
        workbook = excel.Workbooks.Open(archivo)
           # Set the password to protect the file
        workbook.Password = password
        workbook.Save()
        workbook.Close()

    # Cerrar Excel
    excel.Quit()

# Función para abrir todos los archivos Excel en un directorio
def abrir_excel(directorio, password):
    # Obtener la lista de archivos Excel en el directorio
    archivos_excel = glob.glob(os.path.join(directorio, '*.xlsx'))

    # Crear una instancia de Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    # Iterar sobre cada archivo Excel y abrirlo
    for archivo in archivos_excel:
        print(archivo)
        print(password)
        workbook = excel.Workbooks.Open(archivo,Password=password)
        # Set the password to protect the file
        print(password)
        excel.Visible = True

    # Cerrar Excel
    # excel.Quit()

# Definir la interfaz de usuario
layout = [
    [sg.Text('Encriptar archivos Excel')],
    [sg.FolderBrowse(key='directorio'), sg.Text('Directorio:', size=(20, 1))],
    [sg.Text('Contraseña:', size=(10, 1)), sg.InputText(key='password')],
    [sg.Submit(button_text='Encriptar')],
    [sg.Submit(button_text='Abrir')],
    [sg.Text('')],
    [sg.Submit(button_text='Salir',button_color=('red'))]
]

# Crear la ventana de la aplicación
window = sg.Window('Mi aplicación', layout, size=(300, 300))

# Bucle principal de la aplicación
while True:
    event, values = window.Read()

    # Verificar si se ha hecho clic en el botón de salida
    if event == 'Salir' or event == sg.WIN_CLOSED:
        break

    if event == 'Encriptar':
        directorio = values['directorio']
        password = values['password']
        
        # Verificar si el directorio existe y contiene archivos .xlsx
        if not os.path.exists(directorio):
            sg.popup('El directorio especificado no existe.')
        elif not glob.glob(os.path.join(directorio, '*.xlsx')):
            sg.popup('El directorio no contiene archivos Excel (.xlsx).')
        elif not password:
            sg.popup('La contraseña no puede estar vacía.') 
        else:
            encriptar_excel(directorio, password)
            sg.popup('Los archivos han sido encriptados con éxito.')

    if event == 'Abrir':
        directorio = values['directorio']
        password = values['password']
        
        # Verificar si el directorio existe y contiene archivos .xlsx
        if not os.path.exists(directorio):
            sg.popup('El directorio especificado no existe.')
        elif not glob.glob(os.path.join(directorio, '*.xlsx')):
            sg.popup('El directorio no contiene archivos Excel (.xlsx).')
        elif not password:
            sg.popup('La contraseña no puede estar vacía.') 
        else:
            abrir_excel(directorio,password)
            sg.popup('Los archivos han sido abiertos con éxito.')
# Cerrar la ventana de la aplicación
window.Close()
