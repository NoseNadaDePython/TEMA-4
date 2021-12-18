import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')



EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Agregar mas informacion:')],
    [sg.Text('Nombre del libro', size=(15,1)), sg.InputText(key='Nombre del libro')],
    [sg.Text('Formato', size=(15,1)), sg.InputText(key='Formato')],
    [sg.Text('Proveedor', size=(15,1)),sg.InputText(key="Proveedor")],
    [sg.Text('Existencias', size=(15,1)),sg.InputText(key="Existencia")],
    
    
    [sg.Submit(), sg.Button('Clear'),sg.Exit()]
    
]

window = sg.Window('Añadir informacion', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
        
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
    
window.close()