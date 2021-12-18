
from tkinter import Entry, Label, Frame,messagebox, Tk, Button,ttk, Scrollbar, VERTICAL, HORIZONTAL,StringVar,END

import pandas as pd
from pandas.core.frame import DataFrame


#solo guarda 2 dato en el excel al momento de crear el .xslx
class Registro(Frame):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)

        self.frame1 = Frame(master)
        self.frame1.grid(columnspan=2, column=0,row=0)
        self.frame2 = Frame(master, bg='navy')
        self.frame2.grid(column=0, row=1)
        self.frame3 = Frame(master)
        self.frame3.grid(rowspan=2, column=1, row=1)

        self.frame4 = Frame(master, bg='black')
        self.frame4.grid(column=0, row=2)
        
        
        self.codigo = StringVar()
        self.nombre_libro = StringVar()
        self.formato = StringVar()
        self.proveedor = StringVar()
        self.existencia = StringVar()
        self.buscar = StringVar()
       # self.nombre_archivo="Registro"
        self.create_wietgs()

    def create_wietgs(self):
        Label(self.frame1, text = 'R E G I S T R O \t D E \t D A T O S',bg='gray22',fg='white', font=('Orbitron',15,'bold')).grid(column=0, row=0)

        Label(self.frame2, text = 'Agregar Nuevos Datos',fg='white', bg ='navy', font=('Rockwell',12,'bold')).grid(columnspan=2, column=0,row=0, pady=5)
        Label(self.frame2, text = 'Codigo',fg='white', bg ='navy', font=('Rockwell',13,'bold')).grid(column=0,row=1, pady=15)
        Label(self.frame2, text = 'Nombre del\nlibro',fg='white', bg ='navy', font=('Rockwell',11,'bold')).grid(column=0,row=2, pady=15)
        Label(self.frame2, text = 'Formato',fg='white', bg ='navy', font=('Rockwell',13,'bold')).grid(column=0,row=3, pady=15)
        Label(self.frame2, text = 'Proveedor', fg='white',bg ='navy', font=('Rockwell',13,'bold')).grid(column=0,row=4, pady=15)
        Label(self.frame2, text = 'Existencia',fg='white', bg ='navy', font=('Rockwell',13,'bold')).grid(column=0,row=5, pady=15)
           
        Entry(self.frame2,textvariable=self.codigo , font=('Arial',12)).grid(column=1,row=1, padx =5)
        Entry(self.frame2,textvariable=self.nombre_libro , font=('Arial',12)).grid(column=1,row=2)
        Entry(self.frame2,textvariable=self.formato , font=('Arial',12)).grid(column=1,row=3)
        Entry(self.frame2,textvariable=self.proveedor , font=('Arial',12)).grid(column=1,row=4)
        Entry(self.frame2,textvariable=self.existencia , font=('Arial',12)).grid(column=1,row=5)

        Label(self.frame4, text = 'Control',fg='white', bg ='black', font=('Rockwell',12,'bold')).grid(columnspan=3, column=0,row=0, pady=1, padx=4)         
        Button(self.frame4,command= self.agregar_datos, text='REGISTRAR', font=('Arial',10,'bold'), bg='magenta2').grid(column=0,row=1, pady=10, padx=4)
        Button(self.frame4,command = self.limpiar_datos, text='LIMPIAR', font=('Arial',10,'bold'), bg='orange red').grid(column=1,row=1, padx=10)        
        Button(self.frame4,command = self.eliminar_fila, text='ELIMINAR', font=('Arial',10,'bold'), bg='yellow').grid(column=2,row=1, padx=4)
        #Button(self.frame4,command = self.buscar_nombre, text='NOMBRE DEL ARCHIVO', font=('Arial',8,'bold'), bg='orange').grid(columnspan=2,column = 1, row=2)
        Entry(self.frame4,textvariable=self.buscar , font=('Arial',12), width=10).grid(column=0,row=2, pady=1, padx=8)
        Button(self.frame4,command=self.guardar_datos, text='Exportar a Excel', font=('Arial',10,'bold'), bg='green2').grid(columnspan=3,column=0,row=3, pady=8)
         

        self.tabla = ttk.Treeview(self.frame3, height=21)
        self.tabla.grid(column=0, row=0)

        ladox = Scrollbar(self.frame3, orient = HORIZONTAL, command= self.tabla.xview)
        ladox.grid(column=0, row = 1, sticky='ew') 
        ladoy = Scrollbar(self.frame3, orient =VERTICAL, command = self.tabla.yview)
        ladoy.grid(column = 1, row = 0, sticky='ns')

        self.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
        

        self.tabla['columns'] = ('Nombre del libro', 'Formato', 'Proveedor', 'Existencia')

        self.tabla.column('#0', minwidth=100, width=120, anchor='center')
        self.tabla.column('Nombre del libro', minwidth=100, width=130 , anchor='center')
        self.tabla.column('Formato', minwidth=100, width=120, anchor='center' )
        self.tabla.column('Proveedor', minwidth=100, width=120 , anchor='center')
        self.tabla.column('Existencia', minwidth=100, width=105, anchor='center')

        self.tabla.heading('#0', text='Codigo', anchor ='center')
        self.tabla.heading('Nombre del libro', text='Nombre del libro', anchor ='center')
        self.tabla.heading('Formato', text='Formato', anchor ='center')
        self.tabla.heading('Proveedor', text='Proveedor', anchor ='center')
        self.tabla.heading('Existencia', text='Existencia', anchor ='center')


        estilo = ttk.Style(self.frame3)
        estilo.theme_use('alt') #  ('clam', 'alt', 'default', 'classic')
        estilo.configure(".",font= ('Helvetica', 12, 'bold'), foreground='red2')        
        estilo.configure("Treeview", font= ('Helvetica', 10, 'bold'), foreground='black',  background='white')
        estilo.map('Treeview',background=[('selected', 'green2')], foreground=[('selected','black')] )

        self.tabla.bind("<<TreeviewSelect>>", self.obtener_fila)  # seleccionar  fila

    
    def guardar_datos(self):
           
            a=self.nombre_libro.get()
            b=self.formato.get()
            c=self.proveedor.get()
            d=self.existencia.get()
            datos = {'Nombre del libro':a,'Formato':b ,'Proveedor':c, 'Existencia':d} 
            #for
           
            nom_excel  = str("Registro.xlsx")	
            df = pd.DataFrame(datos,columns =  ['Nombre del libro', 'Formato', 'Proveedor', 'Existencia'],index=[1]) 
            df.to_excel(nom_excel,index=False)
    
        
        
    def agregar_datos(self):
        self.tabla.get_children()
        codigo = self.codigo.get()
        nombre = self.nombre_libro.get()
        formato1 = self.formato.get()
        proveedor2 = self.proveedor.get()
        existencia1 = self.existencia.get()
        datos = (nombre, formato1, proveedor2, existencia1)
        if codigo and nombre and formato1 and proveedor2 and existencia1:        
            self.tabla.insert('',0, text = codigo, values=datos)
            
            
    def limpiar_datos(self):
        self.tabla.delete(*self.tabla.get_children())
        self.codigo.set('')
        self.nombre_libro.set('')
        self.formato.set('')
        self.proveedor.set('')
        self.existencia.set('')
        


    def eliminar_fila(self):
        fila = self.tabla.selection()
        if len(fila) !=0:        
            self.tabla.delete(fila)
           


    def obtener_fila(self, event):
        current_item = self.tabla.focus()
        if not current_item:
            return
        data = self.tabla.item(current_item)
        self.nombre_borar = data['values'][0]


def main():
    
   
    ventana = Tk()
    ventana.wm_title("'R E G I S T R O \t D E \t D A T O S'")
    ventana.config(bg='gray22')
    ventana.geometry('900x500')
    ventana.resizable(0,0)
    app = Registro(ventana)
    

    app.mainloop()
    

    

if __name__=="__main__":
    main()        