# Registro y busqueda de datos en excel desde una GUI en Tkinter 

from tkinter import Tk, Label, Button,Entry, Frame, END, font,messagebox
from tkinter import   messagebox, filedialog, ttk, Scrollbar, VERTICAL, HORIZONTAL
from tkinter import Tk, Label, Button, Frame,  messagebox, filedialog, ttk, Scrollbar, VERTICAL, HORIZONTAL
import pandas as pd
import PySimpleGUI as sg
#Este si sirve 


#Instanciacion de la ventana y personalizacion
ventana = Tk()
ventana.config(bg='black')
#ventana.minsize(width=600, height=400)
ventana.resizable(0,0)

ventana.title('Registo de libros')
ventana.iconbitmap("biblioteca.ico")

    
# Declaracion de las variables como listas
nombre_libro,formato1,proveedor1,existencia1 = [],[],[],[]

#Metodo para agregar los datos ingresados
def agregar_datos():
    
	nombre_libro.append(ingresa_nombre_lib.get())
	formato1.append(ingresa_formato.get())
	proveedor1.append(ingresa_proveedor.get())
	existencia1.append(ingresa_existencia.get())
	 # para borrar el contenido de los entry una vez se hallan agreado
	ingresa_nombre_lib.delete(0,END)
	ingresa_formato.delete(0,END)
	ingresa_proveedor.delete(0,END)
	ingresa_existencia.delete(0,END)

#Metodo para guardos los datos y  generar hoja excel
def guardar_datos():
  	
	datos = {'Nombre del libro':nombre_libro,'Formato':formato1, 'Proveedor':proveedor1, 'Existencia':existencia1} 
	nom_excel  = str(nombre_archivo.get() +".xlsx")	
	df = pd.DataFrame(datos,columns =  ['Nombre del libro', 'Formato', 'Proveedor', 'Existencia']) 
	df.to_excel(nom_excel,index=False)
	nombre_archivo.delete(0,END)

#Funcion para cerrar la ventana a traves del button "Salir" 
def on_closing():
    if messagebox.askokcancel("Salir "," ¿Quieres salir?"):
        ventana.destroy()

#Frame 1 que contiene las etiquetas y entradas 
#para ingresar los datos 

frame1 = Frame(ventana, bg='gray15')
frame1.grid(column=0, row=0, sticky='nsew')

#*******************************************

#Frame 2 que contiene las etiquetas y entradas 
#para crear y colocarle un nombre al archivo
#excel que se genere 

frame2 = Frame(ventana, bg='gray16')
frame2.grid(column=2, row=0, sticky='nsew')

#Funcion para llamar a traves de un boton
#la ventana para abrir un archivo excel

def mostrar():
	root = Tk()
	root.config(bg='black')
	root.geometry('600x400')
	root.minsize(width=600, height=400)
	root.title('Lunaria Librerias - Mostrar Existencias')

	root.columnconfigure(0, weight = 25)
	root.rowconfigure(0, weight= 25)
	root.columnconfigure(0, weight = 1)
	root.rowconfigure(1, weight= 1)

	frame1 = Frame(root, bg='gray26')
	frame1.grid(column=0,row=0,sticky='nsew')
	frame2 = Frame(root, bg='gray26')
	frame2.grid(column=0,row=1,sticky='nsew')

	frame1.columnconfigure(0, weight = 1)
	frame1.rowconfigure(0, weight= 1)

	frame2.columnconfigure(0, weight = 1)
	frame2.rowconfigure(0, weight= 1)
	frame2.columnconfigure(1, weight = 1)
	frame2.rowconfigure(0, weight= 1)

	frame2.columnconfigure(2, weight = 1)
	frame2.rowconfigure(0, weight= 1)

	frame2.columnconfigure(3, weight = 2)
	frame2.rowconfigure(0, weight= 1)
 
			
	def abrir_archivo():

		archivo = filedialog.askopenfilename(initialdir ='/', 
												title='Selecione archivo', 
												filetype=(('xlsx files', '*.xlsx*'),('All files', '*.*')))
		indica['text'] = archivo


	def datos_excel():

		datos_obtenidos = indica['text']
		try:
			archivoexcel = r'{}'.format(datos_obtenidos)
			

			df = pd.read_excel(archivoexcel)

		except ValueError:
			messagebox.showerror('Informacion', 'Formato incorrecto')
			return None

		except FileNotFoundError:
			messagebox.showerror('Informacion', 'El archivo esta \n malogrado')
			return None

		Limpiar()

		tabla['column'] = list(df.columns)
		tabla['show'] = "headings"  #encabezado
		

		for columna in tabla['column']:
			tabla.heading(columna, text= columna)
		

		df_fila = df.to_numpy().tolist()
		for fila in df_fila:
			tabla.insert('', 'end', values =fila)


	def Limpiar():
		tabla.delete(*tabla.get_children())
	
	def on_closing():
		if messagebox.askokcancel("Salir "," ¿Quieres salir?"):
			root.destroy()
		   
	tabla = ttk.Treeview(frame1 , height=10)
	tabla.grid(column=0, row=0, sticky='nsew')

	ladox = Scrollbar(frame1, orient = HORIZONTAL, command= tabla.xview)
	ladox.grid(column=0, row = 1, sticky='ew') 

	ladoy = Scrollbar(frame1, orient =VERTICAL, command = tabla.yview)
	ladoy.grid(column = 1, row = 0, sticky='ns')

	tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

	estilo = ttk.Style(frame1)
	estilo.theme_use('clam') #  ('clam', 'alt', 'default', 'classic')
	estilo.configure(".",font= ('Arial', 14), foreground='red2')
	estilo.configure("Treeview", font= ('Helvetica', 12), foreground='black',  background='white')
	estilo.map('Treeview',background=[('selected', 'green2')], foreground=[('selected','black')] )


	boton1 = Button(frame2, text= 'Abrir', bg='green2', command= abrir_archivo)
	boton1.grid(column = 0, row = 0, sticky='nsew', padx=10, pady=10)

	boton2 = Button(frame2, text= 'Mostrar', bg='magenta', command= datos_excel)
	boton2.grid(column = 1, row = 0, sticky='nsew', padx=10, pady=10)

	#salir=Button(frame2,text="Salir",bg="red",command=on_closing)
	#salir.grid(column=1,row=1,sticky="nsew")

	boton3 = Button(frame2, text= 'Salir', bg='red', command=on_closing)
	boton3.grid(column = 2, row = 0, sticky='nsew', padx=10, pady=10)


	indica = Label(frame2, fg= 'white', bg='gray26', text= 'Ubicación del archivo', font= ('Arial',10,'bold') )
	indica.grid(column=3, row = 0)

	root.mainloop()
########################################

#########################################
#"""
def add_data():

	EXCEL_FILE = str(nombre_archivo.get() +".xlsx")	
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

		
		#"""
#############
nombre = Label(frame1, text ='Nombre del\n libro', width=10)
nombre.config(background="gray15",foreground="blue",font=("Arial black",14))
nombre.grid(column=0, row=0, pady=20, padx= 10)
ingresa_nombre_lib = Entry(frame1,  width=20, font = ('Arial',12))
ingresa_nombre_lib.grid(column=1, row=0)

fomarto2 = Label(frame1, text ='Formato', width=10)
fomarto2.config(background="gray15",foreground="blue",font=("Arial black",14))
fomarto2.grid(column=0, row=1, pady=20, padx= 10)
ingresa_formato = Entry(frame1, width=20, font = ('Arial',12))
ingresa_formato.grid(column=1, row=1)

provee = Label(frame1, text ='Proveedor', width=10)
provee.config(background="gray15",foreground="blue",font=("Arial black",14))
provee.grid(column=0, row=2, pady=20, padx= 10)
ingresa_proveedor = Entry(frame1,  width=20, font = ('Arial',12))
ingresa_proveedor.grid(column=1, row=2)

exis = Label(frame1, text ='Existencia', width=10)
exis.config(background="gray15",foreground="blue",font=("Arial black",14))
exis.grid(column=0, row=3, pady=20, padx= 10)
ingresa_existencia = Entry(frame1,  width=20, font = ('Arial',12))
ingresa_existencia.grid(column=1, row=3)

agregar = Button(frame1, width=20, font = ('Arial',12, 'bold'), text='Agregar', bg='orange',bd=5, command =agregar_datos)
agregar.grid(columnspan=2, row=5, pady=20, padx= 10)

archivo = Label(frame2, text ='Ingrese Nombre del archivo', width=25, bg='gray16',font = ('Arial',12, 'bold'), fg='white')
archivo.grid(column=0, row=0, pady=10, padx= 10)

nombre_archivo = Entry(frame2, width=23, font = ('Arial',12),highlightbackground = "green", highlightthickness=4)
nombre_archivo.grid(column=0, row=1, pady=1, padx= 10)

guardar = Button(frame2, width=20, font = ('Arial',12, 'bold'), text='Guardar', bg='green2',bd=5, command =guardar_datos)
guardar.grid(column=0, row=2, pady=20, padx= 10)

salir=Button(frame2,text="Salir",font = ('Arial',12, 'bold'), bg='red',bd=5, command =on_closing)
salir.grid(column=0, row=5, pady=20, padx= 10)

busque=Button(frame2,text="Abrir archivo",command=mostrar,font=("Arial",12,"bold"),bd=5,bg='yellow')
busque.grid(column=0,row=3,padx=15,pady=5)

add=Button(frame2,text='Agregar más datos',font=("Arial",12,'bold'),bd=5,command=add_data)
add.grid(column=0,row=4)


ventana.mainloop()
