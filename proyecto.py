
#actualizacion 23/05/2022
from cProfile import label
from cgitb import text
from faulthandler import disable
from mimetypes import init
from multiprocessing.sharedctypes import Value
from pydoc import doc
from struct import pack
from tkinter import LEFT, RIGHT, SOLID, TOP, Button, Canvas, DoubleVar, Entry, Frame, Image, Label, LabelFrame, OptionMenu, Scrollbar, StringVar, Tk, Toplevel, ttk
from tokenize import Double, String
from turtle import st, width
from docxtpl import DocxTemplate, InlineImage
from tkinter import filedialog 
from PIL import Image, ImageTk
from docx.shared import Mm
from tkinter import messagebox
from tkcalendar import *
from time import sleep
from tkinterdnd2 import DND_FILES,TkinterDnD
import tkinter as tk

#pantalla de carga
class LoadingSplash:
    def __init__(self):
        # setting root window:
        self.secc = Tk()
        self.secc.config(bg="white")
        self.secc.title("Medidor CVM")
        self.secc.iconbitmap("recursos/logo.ico")
        self.secc.geometry("850x500")
        #self.root.attributes("-fullscreen",True)
        imagen_carga= ImageTk.PhotoImage(Image.open('recursos/pan_inicio.png').resize((800, 400)))
        Label(self.secc,image= imagen_carga).place(x=20,y=20)
        # loading text:
        Label(self.secc, text="Loading...", font="Bahnschrift 15",
            bg="white", fg="black").place(x=250, y=430)
        
        # loading blocks:
        for i in range(16):
            Label(self.secc, bg="#1F2732", width=2, height=1).place(x=(i+12)*22, y=460)
        
        # update root to see animation:
        self.secc.update()
        self.play_animation()
        # window in mainloop:
        #self.secc.mainloop()
    # loader animation:
    def play_animation(self):
        for i in range(4):
            for j in range(16):
                # make block yellow:
                Label(self.secc, bg="#FFBD09", width=2, height=1).place(x=(j+12)*22, y=460)
                sleep (0.06)
                self.secc.update_idletasks()
                # make block dark:
                Label(self.secc, bg="#1F2732", width=2, height=1).place(x=(j+12)*22, y=460)
        else:
            self.secc.destroy()
            #exit(0)
for i in range(2):
    print("iterado :",i)
    if i==0 :
        LoadingSplash() 
    if i==1:

        #interfaz
        raiz=TkinterDnD.Tk()
        raiz.title("Medidor CVM")
        raiz.iconbitmap("recursos/logo.ico")

        #datos de la interfaz
        kkkinforme_tecnico = StringVar()
        kkkfecha_lectura = StringVar()
        kkkfecha_lectura.set("dd/mm/yyyy")
        kkkfecha_emision = StringVar()
        kkkfecha_emision.set("dd/mm/yyyy")
        kkkcliente = StringVar()
        kkkarea = StringVar()
        kkkdias=StringVar()
        kkkdias.set("0")

        nombre_meses = [
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre"
        ]
        kkk_nom_mes= StringVar()
        kkk_nom_mes.set(nombre_meses[0])
        kkkmes1 = StringVar()
        kkkmes1.set("0")
        kkknom_mes1 = StringVar()
        kkknom_mes1.set("Mes")
        kkkmes2 = StringVar()
        kkkmes2.set("0")
        kkknom_mes2 = StringVar()
        kkknom_mes2.set("Mes")
        kkkmes3 = StringVar()
        kkkmes3.set("0")
        kkknom_mes3 = StringVar()
        kkknom_mes3.set("Mes")
        kkkmes4 = StringVar()
        kkkmes4.set("0")
        kkknom_mes4 = StringVar()
        kkknom_mes4.set("Mes")
        kkkmes5 = StringVar()
        kkkmes5.set("0")
        kkknom_mes5 = StringVar()
        kkknom_mes5.set("Mes")
        kkkmes6 = StringVar()
        kkkmes6.set("0")
        kkknom_mes6 = StringVar()
        kkknom_mes6.set("Mes")

        kkkcargo_fijo_mensual = StringVar()
        kkkcargo_fijo_mensual.set("0")
        kkkcargo_energia_activa_punta = StringVar()
        kkkcargo_energia_activa_punta.set("0")
        kkkcargo_energia_activa_fuera_punta = StringVar()
        kkkcargo_energia_activa_fuera_punta.set("0")
        kkkcargo_potencia_activa_generacion_presente_punta = StringVar()
        kkkcargo_potencia_activa_generacion_presente_punta.set("0")
        kkkcargo_potencia_activa_generacion_presente_fuera_punta = StringVar()
        kkkcargo_potencia_activa_generacion_presente_fuera_punta.set("0")
        kkkcargo_potencia_activa_redes_presente_punta = StringVar()
        kkkcargo_potencia_activa_redes_presente_punta.set("0")
        kkkcargo_potencia_activa_redes_presente_fuera_punta = StringVar()
        kkkcargo_potencia_activa_redes_presente_fuera_punta.set("0")
        kkkcargo_energia_reactiva_exc_30 = StringVar()
        kkkcargo_energia_reactiva_exc_30.set("0")

        kkkcont_tableros=StringVar()
        kkkcont_tableros.set("0")
        kkknom_tablero=StringVar()

        kkkfoto1=StringVar()
        kkkfoto1=""
        kkkfoto2=StringVar()
        kkkfoto2=""
        kkkfoto3=StringVar()
        kkkfoto3=""

        kkkmost_foto1=StringVar()
        kkkmost_foto2=StringVar()
        kkkmost_foto3=StringVar()

        kkkfp_mes_actual =  StringVar()
        kkkfp_mes_actual.set("0")
        kkkfp_mes_anterior =  StringVar()
        kkkfp_mes_anterior.set("0")

        kkkhp_mes_actual =  StringVar()
        kkkhp_mes_actual.set("0")
        kkkhp_mes_anterior =  StringVar()
        kkkhp_mes_anterior.set("0")

        kkkea_mes_actual =  StringVar()
        kkkea_mes_actual.set("0")
        kkkea_mes_anterior = StringVar()
        kkkea_mes_anterior.set("0")

        kkkmaxima_demanda = StringVar()
        kkkmaxima_demanda.set("0") 

        kkker_mes_actual =  StringVar()
        kkker_mes_actual.set("0")
        kkker_mes_anterior =  StringVar()
        kkker_mes_anterior.set("0")

        icono_check= ImageTk.PhotoImage(Image.open('recursos/check.png').resize((20, 20)))

        #datos donde se guarda los datos de la interfaz

        informe_tecnico = StringVar()
        nom_tablero=[]
        tablero = ""
        fecha_lectura = StringVar()
        fecha_emision = StringVar()
        cliente = StringVar()
        area = StringVar()
        dias= 0

        maxima_demanda_del_mes=0
        nombre_mes = []
        meses = []
        promedio=0.0

        cargo_fijo_mensual = 0.0
        cargo_energia_activa_punta = 0.0
        cargo_energia_activa_fuera_punta = 0.0
        cargo_potencia_activa_generacion_presente_punta = 0.0
        cargo_potencia_activa_generacion_presente_fuera_punta = 0.0
        cargo_potencia_activa_redes_presente_punta = 0.0
        cargo_potencia_activa_redes_presente_fuera_punta = 0.0
        cargo_energia_reactiva_exc_30 = 0.0

        cont_tableros=0

        foto1=[]
        foto2=[]
        foto3=[]

        fp_mes_actual =  []
        fp_mes_anterior =  []
        fp_mes =  []
        hp_mes_actual =  [] 
        hp_mes_anterior =  [] 
        hp_mes =  [] 
        ea_mes_actual =  [] 
        ea_mes_anterior = [] 
        ea_mes =  [] 
        maxima_demanda =  [] 
        er_mes_actual =  [] 
        er_mes_anterior =  [] 
        er_mes = []

        nom_mes1 = ""
        nom_mes2 = ""
        nom_mes3 = ""
        nom_mes4 = ""
        nom_mes5 = ""
        nom_mes6 = ""

        fp_mestotal = 0
        hp_mestotal = 0
        ea_mestotal = 0
        maxima_demandatotal = 0
        er_mestotal =  0

        cali_tari= 0
        horas_punta = 0

        #datos solo para el rellenado del word
        texto=""
        cargo_eapp = 0.0
        cargo_eafp = 0.0
        cargo_pagpp = 0.0
        cargo_pagfp = 0.0
        cargo_parpp = 0.0
        cargo_parfp = 0.0
        operacion1 =""
        operacion2 =""
        operacion3 =""
        operacion4 =""
        v2= 0.0
        er_mestotalfac= 0.0
        texto2= ""
        v3= 0.0
        cargo_er_exc30 =0.0
        v4= 0
        v5= 0
        v6= 0
        v7= 0
        subtotal= 0.0
        conigv= 0.0
        total_final= 0
        anio=""
        p1=0
        p2=0
        p3=0

        def drop_inside_list_box(event):
            listb.insert("end",event.data)


        def formato_fecha(fecha,v):
            global fecha_lectura
            global fecha_emision
            global anio
            mes=""
            partes_fecha = fecha.split('-')
            if partes_fecha[1]=="01":
                mes="Enero"
            if partes_fecha[1]=="02":
                mes="Febrero"
            if partes_fecha[1]=="03":
                mes="Marzo"
            if partes_fecha[1]=="04":
                mes="Abril"
            if partes_fecha[1]=="05":
                mes="Mayo"
            if partes_fecha[1]=="06":
                mes="Junio"
            if partes_fecha[1]=="07":
                mes="Julio"
            if partes_fecha[1]=="08":
                mes="Agosto"
            if partes_fecha[1]=="09":
                mes="Septiembre"
            if partes_fecha[1]=="10":
                mes="Octubre"
            if partes_fecha[1]=="11":
                mes="Noviembre"
            if partes_fecha[1]=="12":
                mes="Diciembre"
            
            if v==1:
                fecha_lectura=str('{}{}{}{}{}'.format(partes_fecha[0],' de ',mes,' del ',partes_fecha[2]))
                fecha_lectura=fecha_lectura.upper()
            if v==2:
                fecha_emision=str('{}{}{}{}{}'.format(partes_fecha[0],' de ',mes,' del ',partes_fecha[2]))
                anio=str(partes_fecha[2])

        def mensaje1():
            messagebox.showwarning('Mensaje de advertencia','Rellene todos los campos de la seccion "Tableros"')

        def mensaje2():
            messagebox.showwarning('Mensaje de advertencia','Rellene todos los campos con datos reales, asegurese de tener por lo menos 1 medidor agregado')

        def mensaje3(nombre):
            messagebox.showinfo('Mensaje informativo',nombre+' a sido creado exitosamente')

        def reniciar_arreglo(arreglo):
                for x in range(len(arreglo)-1,-1,-1):
                        arreglo.pop(x)

        def reiniciar():
                global kkkinforme_tecnico 
                global kkkfecha_lectura
                global kkkfecha_emision
                global kkkcliente
                global kkkarea
                global kkkdias
                global kkkmes1 
                global kkkmes2 
                global kkkmes3 
                global kkkmes4 
                global kkkmes5 
                global kkkmes6 
                global kkkcargo_fijo_mensual 
                global kkkcargo_energia_activa_punta 
                global kkkcargo_energia_activa_fuera_punta 
                global kkkcargo_potencia_activa_generacion_presente_punta 
                global kkkcargo_potencia_activa_generacion_presente_fuera_punta 
                global kkkcargo_potencia_activa_redes_presente_punta 
                global kkkcargo_potencia_activa_redes_presente_fuera_punta 
                global kkkcargo_energia_reactiva_exc_30 
                global kkkcont_tableros
                global kkkfoto1
                global kkkfoto2
                global kkkfoto3
                global kkkmost_foto1
                global kkkmost_foto2
                global kkkmost_foto3
                global kkkfp_mes_actual 
                global kkkfp_mes_anterior 
                global kkkhp_mes_actual 
                global kkkhp_mes_anterior 
                global kkkea_mes_actual 
                global kkkea_mes_anterior 
                global kkkmaxima_demanda  
                global kkker_mes_actual 
                global kkker_mes_anterior
                global informe_tecnico 
                global tablero 
                global fecha_lectura 
                global fecha_emision
                global cliente 
                global area 
                global dias
                global meses 
                global maxima_demanda_del_mes
                global promedio
                global cargo_fijo_mensual 
                global cargo_energia_activa_punta 
                global cargo_energia_activa_fuera_punta 
                global cargo_potencia_activa_generacion_presente_punta 
                global cargo_potencia_activa_generacion_presente_fuera_punta 
                global cargo_potencia_activa_redes_presente_punta 
                global cargo_potencia_activa_redes_presente_fuera_punta 
                global cargo_energia_reactiva_exc_30 
                global cont_tableros
                global foto1
                global foto2
                global foto3
                global fp_mes_actual 
                global fp_mes_anterior 
                global fp_mes 
                global hp_mes_actual  
                global hp_mes_anterior 
                global hp_mes 
                global ea_mes_actual  
                global ea_mes_anterior  
                global ea_mes 
                global maxima_demanda  
                global er_mes_actual  
                global er_mes_anterior  
                global er_mes 
                global fp_mestotal 
                global hp_mestotal 
                global ea_mestotal 
                global maxima_demandatotal 
                global er_mestotal 
                global cali_tari
                global horas_punta 
                global texto
                global cargo_eapp 
                global cargo_eafp 
                global cargo_pagpp 
                global cargo_pagfp 
                global cargo_parpp 
                global cargo_parfp 
                global operacion1
                global operacion2 
                global operacion3 
                global operacion4 
                global v2
                global er_mestotalfac
                global texto2
                global v3
                global cargo_er_exc30 
                global v4
                global v5
                global v6
                global v7
                global subtotal
                global conigv
                global total_final
                global kkknom_tablero
                global nom_tablero
                global estadolabel
                
                kkkinforme_tecnico.set("")
                kkkfecha_lectura.set("dd-mm-yyyy")
                kkkfecha_emision.set("dd-mm-yyyy")
                kkkcliente.set("")
                kkkarea.set("")
                kkkdias.set("0")
                kkkmes1.set("0")
                kkkmes2.set("0")
                kkkmes3.set("0")
                kkkmes4.set("0")
                kkkmes5.set("0")
                kkkmes6.set("0")
                kkknom_tablero.set("")
                kkkfp_mes_actual.set("0")
                kkkfp_mes_anterior.set("0")
                kkkhp_mes_actual.set("0")
                kkkhp_mes_anterior.set("0")
                kkkea_mes_actual.set("0")
                kkkea_mes_anterior.set("0")
                kkkmaxima_demanda.set("0")
                kkker_mes_actual.set("0")
                kkker_mes_anterior.set("0")
                kkkmost_foto1.set("")
                kkkmost_foto2.set("")
                kkkmost_foto3.set("")
                kkkcont_tableros.set("0")
                maxima_demanda_del_mes=0

                reniciar_arreglo(nom_tablero)
                reniciar_arreglo(meses)
                reniciar_arreglo(foto1)
                reniciar_arreglo(foto2)
                reniciar_arreglo(foto3)
                reniciar_arreglo(fp_mes_actual)
                reniciar_arreglo(fp_mes_anterior)
                reniciar_arreglo(fp_mes)
                reniciar_arreglo(hp_mes_actual)
                reniciar_arreglo(hp_mes_anterior)
                reniciar_arreglo(hp_mes)
                reniciar_arreglo(ea_mes_actual)
                reniciar_arreglo(ea_mes_anterior)
                reniciar_arreglo(ea_mes)
                reniciar_arreglo(maxima_demanda)
                reniciar_arreglo(er_mes_actual)
                reniciar_arreglo(er_mes_anterior)
                reniciar_arreglo(er_mes)
                cont_tableros=0
                fp_mestotal= 0
                hp_mestotal= 0
                ea_mestotal= 0
                maxima_demandatotal= 0
                er_mestotal= 0
                btnagregar.configure(state="normal")
                estadolabel=Label(miframe, text="Estado: falta guardar        ",font=("Arial", 10)).place(x=10,y=760)

        def cambiar_nombre_mes():
            band = 0
            if kkk_nom_mes.get()=="Diciembre":
                band = 11
            if kkk_nom_mes.get()=="Noviembre":
                band = 10
            if kkk_nom_mes.get()=="Octubre":
                band = 9
            if kkk_nom_mes.get()=="Septiembre":
                band = 8
            if kkk_nom_mes.get()=="Agosto":
                band = 7
            if kkk_nom_mes.get()=="Julio":
                band = 6
            if kkk_nom_mes.get()=="Junio":
                band = 5
            if kkk_nom_mes.get()=="Mayo":
                band = 4
            if kkk_nom_mes.get()=="Abril":
                band = 3
            if kkk_nom_mes.get()=="Marzo":
                band = 2
            if kkk_nom_mes.get()=="Febrero":
                band = 1
            if kkk_nom_mes.get()=="Enero":
                band = 0
            
            kkknom_mes6.set(nombre_meses[band])
            kkknom_mes5.set(nombre_meses[band-1])
            kkknom_mes4.set(nombre_meses[band-2])
            kkknom_mes3.set(nombre_meses[band-3])
            kkknom_mes2.set(nombre_meses[band-4])
            kkknom_mes1.set(nombre_meses[band-5])

        wraper1 = LabelFrame(raiz,width=0)

        mycanvas = Canvas(wraper1,width=920,height=1500)
        mycanvas.pack(side=LEFT,expand=1)

        yscrollbar=ttk.Scrollbar(wraper1,orient="vertical", command=mycanvas.yview)
        yscrollbar.pack(side=RIGHT,fill="y",expand=1)

        mycanvas.configure(yscrollcommand=yscrollbar.set)

        mycanvas.bind('<Configure>',lambda e: mycanvas.configure(scrollregion = mycanvas.bbox('all')))

        #Empieza la interfaz
        miframe= Frame(mycanvas,width=1200,height=800)
        miframe.pack(side=TOP)
        miframe.config(bg="#ecefee")

        #titulo
        frametitulo= Frame(miframe,width=1200,height=100)
        frametitulo.pack()
        frametitulo.config(bg="#ecefee")

        grantitulolabel=Label(frametitulo, text="MEDIDOR CVM",font=("Arial Black", 22))
        grantitulolabel.grid(row=0, column=0, pady=2,padx=5)
        grantitulolabel.config(bg="#ecefee")

        #cabecera
        framecabecera= Frame(miframe,width=1200,height=100,bd=1,relief="solid")
        framecabecera.pack()
        #framecabecera.config(bg="black")

        #framecabe1
        framecabe1= Frame(framecabecera,bd=1,relief="solid")
        framecabe1.grid(row=0,column=0,sticky='w')
        framecabe1.config(bg="#ecefee",pady=0,padx=0)
        ####
        titulo800label=Label(framecabe1, text="DATOS  GENERALES",font=("Arial Black", 12))
        titulo800label.grid(row=0, column=0, pady=5,padx=360,columnspan=6)

        izquierda1= Frame(framecabe1)
        izquierda1.grid(row=1,column=0,padx=25)

        centro1= Frame(framecabe1)
        centro1.grid(row=1,column=1,padx=25)

        derecha1= Frame(framecabe1)
        derecha1.grid(row=1,column=2,padx=25)

        fecha_doclabel=Label(derecha1, text="Informe Técnico:",font=("Arial", 9))
        fecha_doclabel.grid(row=1, column=4, pady=5,padx=0,sticky="e")
        cuadroinforme_tecnico=Entry(derecha1,textvariable=kkkinforme_tecnico,font=("Arial", 9))
        cuadroinforme_tecnico.grid(row=1, column=5, pady=5,padx=10,sticky="w")

        fecha_lecturalabel=Label(centro1, text="Fecha de Lectura:",font=("Arial", 9))
        fecha_lecturalabel.grid(row=1, column=2, sticky="e", pady=5,padx=0)
        cuadrofecha_lectura=DateEntry(centro1,textvariable=kkkfecha_lectura,year=2022,date_pattern="dd-mm-yyyy",locale="es",font=("Arial", 9))
        cuadrofecha_lectura.grid(row=1, column=3, sticky="w", pady=5,padx=10)

        fecha_emisionlabel=Label(centro1, text="Fecha de Emision:",font=("Arial", 9))
        fecha_emisionlabel.grid(row=2, column=2, sticky="e", pady=5,padx=0)
        cuadrofecha_emision=DateEntry(centro1,textvariable=kkkfecha_emision,year=2022,date_pattern="dd-mm-yyyy",locale="es",font=("Arial", 9))
        cuadrofecha_emision.grid(row=2, column=3, sticky="w", pady=5,padx=10)

        clientelabel=Label(izquierda1, text="Cliente:",font=("Arial", 9))
        clientelabel.grid(row=0, column=0, sticky="e", pady=5, padx=0)
        cuadrocliente=Entry(izquierda1,textvariable=kkkcliente,font=("Arial", 9))
        cuadrocliente.grid(row=0, column=1, sticky="w", pady=5,padx=10)

        arealabel=Label(izquierda1, text="Area:",font=("Arial", 9))
        arealabel.grid(row=1, column=0, sticky="e", pady=5, padx=0)
        cuadroarea=Entry(izquierda1,textvariable=kkkarea,font=("Arial", 9))
        cuadroarea.grid(row=1, column=1, sticky="w", pady=5,padx=10)

        diaslabel=Label(derecha1, text="Dias del mes:",font=("Arial", 9))
        diaslabel.grid(row=2, column=4, sticky="e", pady=5, padx=0)
        cuadrodias=Entry(derecha1,textvariable=kkkdias,width=8,font=("Arial", 9))
        cuadrodias.grid(row=2, column=5, sticky="w", pady=5,padx=10)

        #framecabe2
        framecabe2= Frame(framecabecera,bd=1,relief="solid")
        framecabe2.grid(row=1,column=0)
        framecabe2.config(bg="#ecefee",padx=100,pady=4)

        titulo80label=Label(framecabe2, text="MAXIMA DEMANDA ",font=("Arial Black", 12))
        titulo80label.grid(row=0, column=0, pady=5,padx=265,columnspan=6)

        #frame select
        frameselect= Frame(framecabe2,relief="solid")
        frameselect.grid(row=1,column=0,columnspan=6)
        frameselect.config(bg="#ecefee",padx=50,pady=5)

        #elegirmeslabel=Label(frameselect, text="Mes actual:",font=("Arial", 9))
        #elegirmeslabel.grid(row=0, column=0, sticky="e")

        opc_mes = OptionMenu (frameselect, kkk_nom_mes, *nombre_meses)
        opc_mes.grid(row=0, column=1)
        opc_mes.config(width=12)

        boton_check = Button(frameselect,image=icono_check, command=cambiar_nombre_mes)
        boton_check.grid(row=0,column=2)

        izquierda2= Frame(framecabe2)
        izquierda2.grid(row=2,column=0,padx=25)

        centro2= Frame(framecabe2)
        centro2.grid(row=2,column=1,padx=25)

        derecha2= Frame(framecabe2)
        derecha2.grid(row=2,column=2,padx=25)

        entrymes1=Entry(izquierda2 , textvariable=kkknom_mes1,width=11)
        entrymes1.configure(state='disabled')
        entrymes1.grid(row=0, column=0, sticky="e", pady=5,padx=10)
        cuadromes1=Entry(izquierda2, textvariable=kkkmes1,width=8)
        cuadromes1.grid(row=0, column=1, sticky="w", pady=5,padx=0)
        taaa1label=Label(izquierda2, text="(kW)",font=("Arial", 9))
        taaa1label.grid(row=0, column=2, sticky="w")

        entrymes2=Entry(centro2 , textvariable=kkknom_mes2,width=11)
        entrymes2.configure(state='disabled')
        entrymes2.grid(row=0, column=0, sticky="e", pady=5,padx=10)
        cuadromes2=Entry(centro2, textvariable=kkkmes2,width=8)
        cuadromes2.grid(row=0, column=1, sticky="w", pady=5,padx=0)
        taaa2label=Label(centro2, text="(kW)",font=("Arial", 9))
        taaa2label.grid(row=0, column=2, sticky="w")

        entrymes3=Entry(derecha2 , textvariable=kkknom_mes3,width=11)
        entrymes3.configure(state='disabled')
        entrymes3.grid(row=0, column=0, sticky="e", pady=5,padx=10)
        cuadromes3=Entry(derecha2 , textvariable=kkkmes3,width=8)
        cuadromes3.grid(row=0, column=1, sticky="w", pady=5,padx=0)
        taaa3label=Label(derecha2, text="(kW)",font=("Arial", 9))
        taaa3label.grid(row=0, column=2, sticky="w")

        entrymes4=Entry(izquierda2 , textvariable=kkknom_mes4,width=11)
        entrymes4.configure(state='disabled')
        entrymes4.grid(row=1, column=0, sticky="e", pady=5,padx=10)
        cuadromes4=Entry(izquierda2, textvariable=kkkmes4,width=8)
        cuadromes4.grid(row=1, column=1, sticky="w", pady=5,padx=0)
        taaa4label=Label(izquierda2, text="(kW)",font=("Arial", 9))
        taaa4label.grid(row=1, column=2, sticky="w")

        entrymes5=Entry(centro2 , textvariable=kkknom_mes5,width=11)
        entrymes5.configure(state='disabled')
        entrymes5.grid(row=1, column=0, sticky="e", pady=5,padx=10)
        cuadromes5=Entry(centro2 , textvariable=kkkmes5,width=8)
        cuadromes5.grid(row=1, column=1, sticky="w", pady=5,padx=0)
        taaa5label=Label(centro2, text="(kW)",font=("Arial", 9))
        taaa5label.grid(row=1, column=2, sticky="w")

        entrymes6=Entry(derecha2 , textvariable=kkknom_mes6,width=11)
        entrymes6.configure(state='disabled')
        entrymes6.grid(row=1, column=0, sticky="e", pady=5,padx=10)
        cuadromes6=Entry(derecha2 , textvariable=kkkmes6,width=8)
        cuadromes6.configure(state='disabled')
        cuadromes6.grid(row=1, column=1, sticky="w", pady=5,padx=0)
        taaa6label=Label(derecha2, text="(kW)",font=("Arial", 9))
        taaa6label.grid(row=1, column=2, sticky="w")

        #tarifario
        frametarifario= Frame(miframe,width=1200,height=151,bd=2,relief="solid")
        frametarifario.pack()
        frametarifario.config(bg="#ecefee", padx=5)

        titulo2label=Label(frametarifario, text="PLIEGO TARIFARIO",font=("Arial Black", 12))
        titulo2label.grid(row=0, column=0, columnspan=8,pady=5,padx=0)

        ta0label=Label(frametarifario, text="C.Fijo Mensual\n(S/./mes):",font=("Arial", 9))
        ta0label.grid(row=2, column=0, sticky="e", pady=5)
        cuadrota0=Entry(frametarifario,textvariable=kkkcargo_fijo_mensual,width=10)
        cuadrota0.grid(row=2, column=1, sticky="w", pady=5,padx=12)

        ta1label=Label(frametarifario, text="C.Energía Activa Punta\n(Ctm.S/./kW.h):",font=("Arial", 9))
        ta1label.grid(row=2, column=2, sticky="e", pady=5)
        cuadrota1=Entry(frametarifario,textvariable=kkkcargo_energia_activa_punta,width=10)
        cuadrota1.grid(row=2, column=3, sticky="w", pady=5,padx=12)

        ta2label=Label(frametarifario, text="C.Energía Activa Fuera\nPunta (Ctm.S/./kW.h):",font=("Arial", 9))
        ta2label.grid(row=4, column=0, sticky="e", pady=5,padx=5)
        cuadrota2=Entry(frametarifario,textvariable=kkkcargo_energia_activa_fuera_punta ,width=10)
        cuadrota2.grid(row=4, column=1, sticky="w", pady=5,padx=12)

        ##
        titulo3label=Label(frametarifario, text="Cargo por Potencia Activa de generación para Usuarios(S/./kW-mes)",font=("Arial", 9))
        titulo3label.grid(row=1, column=4, pady=5,padx=5,columnspan=4)
        titulo3label.config(bg="white")

        ta3label=Label(frametarifario, text="Presente en Punta:",font=("Arial", 9))
        ta3label.grid(row=2, column=4, sticky="e", pady=5)
        cuadrota3=Entry(frametarifario, textvariable=kkkcargo_potencia_activa_generacion_presente_punta ,width=10)
        cuadrota3.grid(row=2, column=5, sticky="w", pady=5,padx=5)

        ta4label=Label(frametarifario, text="Presente Fuera de Punta:",font=("Arial", 9))
        ta4label.grid(row=2, column=6, sticky="e", pady=5)
        cuadrota4=Entry(frametarifario,textvariable=kkkcargo_potencia_activa_generacion_presente_fuera_punta ,width=10)
        cuadrota4.grid(row=2, column=7, sticky="w", pady=5,padx=5)

        ##
        titulo4label=Label(frametarifario, text="Cargo por Potencia Activa de redes de distribución para Usuarios(S/./kW-mes)",font=("Arial", 9))
        titulo4label.grid(row=3, column=4, pady=5,padx=5,columnspan=4)
        titulo4label.config(bg="white")

        ta5label=Label(frametarifario, text="Presente en Punta:",font=("Arial", 9))
        ta5label.grid(row=4, column=4, sticky="e", pady=5)
        cuadrota5=Entry(frametarifario, textvariable=kkkcargo_potencia_activa_redes_presente_punta ,width=10)
        cuadrota5.grid(row=4, column=5, sticky="w", pady=5,padx=5)

        ta6label=Label(frametarifario, text="Presente Fuera de Punta:",font=("Arial", 9))
        ta6label.grid(row=4, column=6, sticky="e", pady=5)
        cuadrota6=Entry(frametarifario, textvariable=kkkcargo_potencia_activa_redes_presente_fuera_punta ,width=10)
        cuadrota6.grid(row=4, column=7, sticky="w", pady=5,padx=5)

        ta7label=Label(frametarifario, text="C.Energía Reactiva exc.\n30% (Ctm.S/./kVar.h):",font=("Arial", 9))
        ta7label.grid(row=4, column=2, sticky="e", pady=5)
        cuadrota7=Entry(frametarifario,textvariable=kkkcargo_energia_reactiva_exc_30 ,width=10)
        cuadrota7.grid(row=4, column=3, sticky="w", pady=5,padx=12)

        #medidicion
        framemedicion= Frame(miframe,width=1200,height=400,bd=1,relief="solid")
        framemedicion.pack()
        framemedicion.config(bg="#ecefee")

        #tableros
        tablerosss= Frame(framemedicion,bd=1,relief="solid")
        tablerosss.config(bg="#ecefee")
        tablerosss.grid(row=0, column=0 ,columnspan=4,padx=0)

        titulo5label=Label(tablerosss, text="TABLEROS",font=("Arial Black", 12))
        titulo5label.grid(row=0, column=0, pady=5,padx=402 ,columnspan=2)

        titulo84label=Label(tablerosss, text="Nombre del tablero:",font=("Arial", 9))
        titulo84label.grid(row=1, column=0, sticky="e", pady=5,padx=10)
        cuadrota790=Entry(tablerosss,textvariable=kkknom_tablero)
        cuadrota790.grid(row=1, column=1, sticky="w", pady=5,padx=5)

        #ab111
        ab111= Frame(framemedicion,height=150,bd=1,relief="solid")
        ab111.config(bg="#ecefee",padx=15,pady=5)
        ab111.grid(row=1, column=0)

        titulo6label=Label(ab111, text="Energía activa en hora fuera \nde punta, T1 (M/KWh)",font=("Arial", 9))
        titulo6label.grid(row=0, column=0, pady=5,padx=10, columnspan=2)
        titulo6label.config(bg="white")

        ta8label=Label(ab111, text="Mes actual:",font=("Arial", 9))
        ta8label.grid(row=1, column=0, sticky="e", pady=5,padx=2)
        cuadrota8=Entry(ab111, textvariable=kkkfp_mes_actual,width=14)
        cuadrota8.grid(row=1, column=1, sticky="w", pady=5,padx=2)

        ta9label=Label(ab111, text="Mes anterior:",font=("Arial", 9))
        ta9label.grid(row=2, column=0, sticky="e", pady=5,padx=2)
        cuadrota9=Entry(ab111,textvariable=kkkfp_mes_anterior,width=14)
        cuadrota9.grid(row=2, column=1, sticky="w", pady=5,padx=2)

        titulo7label=Label(ab111, text="Energía activa en hora punta,\n T2 (M/KWh)",font=("Arial", 9))
        titulo7label.grid(row=0, column=2, pady=5,padx=10,columnspan=2)
        titulo7label.config(bg="white")

        ta10label=Label(ab111, text="Mes actual:",font=("Arial", 9))
        ta10label.grid(row=1, column=2, sticky="e", pady=5,padx=2)
        cuadrota10=Entry(ab111,textvariable=kkkhp_mes_actual,width=14)
        cuadrota10.grid(row=1, column=3, sticky="w", pady=5,padx=2)

        ta11label=Label(ab111, text="Mes anterior:",font=("Arial", 9))
        ta11label.grid(row=2, column=2, sticky="e", pady=5,padx=2)
        cuadrota11=Entry(ab111,textvariable=kkkhp_mes_anterior,width=14)
        cuadrota11.grid(row=2, column=3, sticky="w", pady=5,padx=2)


        def abreFichero1():
            global kkkfoto1
            imgab = filedialog.askopenfilename(
                parent=ab111,
                initialdir='/Imagenes',
                initialfile='img',
                filetypes=[
                    ("PNG", "*.png"),
                    ("JPEG", "*.jpg"),
                    ("All files", "*")])
            kkkfoto1=imgab
            kkkmost_foto1.set(kkkfoto1)
            print(kkkfoto1)
    

        btnimg1=Button(ab111,text="Abrir imagen", command= abreFichero1,font=("Arial",9))
        btnimg1.grid(row=3,column=0,columnspan=4, pady=5)
        cuadrota500=Entry(ab111, textvariable=kkkmost_foto1,width=25)
        cuadrota500.configure(state='disabled')
        cuadrota500.grid(row=4, column=0, columnspan=4, pady=5)

        #c111
        c111= Frame(framemedicion,height=150,bd=1,relief="solid")
        c111.config(bg="#ecefee",pady=28)
        c111.grid(row=1, column=1,rowspan=7)

        titulo8label=Label(c111, text="Máxima demanda (kW)",font=("Arial", 9))
        titulo8label.grid(row=0, column=0, sticky="n", pady=5,padx=5, columnspan=2)
        titulo8label.config(bg="white")

        ta10label=Label(c111, text="Max:",font=("Arial", 9))
        ta10label.grid(row=1, column=0, sticky="e", pady=5,padx=2)
        cuadrota10=Entry(c111,textvariable=kkkmaxima_demanda,width=14)
        cuadrota10.grid(row=1, column=1, sticky="w", pady=5,padx=2)


        def abreFichero2():
            global kkkfoto2
            imgc = filedialog.askopenfilename(
                parent=c111,
                initialdir='/Imagenes',
                initialfile='img',
                filetypes=[
                    ("PNG", "*.png"),
                    ("JPEG", "*.jpg"),
                    ("All files", "*")])
            kkkfoto2=imgc
            kkkmost_foto2.set(kkkfoto2)


        btnimg2=Button(c111,text="Abrir imagen", command= abreFichero2,font=("Arial",9))
        btnimg2.grid(row=2,column=0,columnspan=2, pady=5)
        cuadrota500=Entry(c111, textvariable=kkkmost_foto2,width=25)
        cuadrota500.configure(state='disabled')
        cuadrota500.grid(row=3, column=0, columnspan=2, pady=5,padx=10)


        #d111
        d111= Frame(framemedicion,height=150,bd=1,relief="solid")
        d111.config(bg="#ecefee",pady=5, padx=17)
        d111.grid(row=1, column=2,rowspan=7)

        titulo9label=Label(d111, text="Energía reactiva inductiva \ntotal (M/KvarLh)",font=("Arial", 9))
        titulo9label.grid(row=0, column=0, pady=5,padx=10, columnspan=2)
        titulo9label.config(bg="white")

        ta12label=Label(d111, text="Mes actual:",font=("Arial",9))
        ta12label.grid(row=1, column=0, sticky="e", pady=5,padx=2)
        cuadrota12=Entry(d111,textvariable=kkker_mes_actual,width=14)
        cuadrota12.grid(row=1, column=1, sticky="w", pady=5,padx=2)

        ta13label=Label(d111, text="Mes anterior:",font=("Arial", 9))
        ta13label.grid(row=2, column=0, sticky="e", pady=5,padx=2)
        cuadrota13=Entry(d111, textvariable=kkker_mes_anterior,width=14)
        cuadrota13.grid(row=2, column=1, sticky="w", pady=5,padx=2)

        def abreFichero3():
            global kkkfoto3
            imgd = filedialog.askopenfilename(
                parent=d111,
                initialdir="/Imagenes",
                initialfile='img',
                filetypes=(
                    ("PNG", "*.png"),
                    ("JPEG", "*.jpg"),
                    ("All files", "*")))
            
            kkkfoto3=imgd
            kkkmost_foto3.set(kkkfoto3)

            print(kkkfoto3)

        btnimg3=Button(d111,text="Abrir imagen", command= abreFichero3,font=("Arial",9))
        btnimg3.grid(row=3,column=0,columnspan=2,pady=5)
        cuadrota500=Entry(d111, textvariable=kkkmost_foto3, width=25)
        cuadrota500.configure(state='disabled')
        cuadrota500.grid(row=4, column=0, columnspan=2, pady=5,padx=10)

        #e111
        e111= Frame(framemedicion,height=150,bd=1,relief="solid")
        e111.config(bg="#ecefee",pady=6,padx=9)
        e111.grid(row=1, column=3, sticky="e",rowspan=7)


        titulo10label=Label(e111, text="Ingresados",font=("Arial", 12))
        titulo10label.grid(row=0, column=0, pady=2,padx=10)
        cuadrota14=Entry(e111, textvariable=kkkcont_tableros,font=("Arial", 50),width=2)
        cuadrota14.configure(state='disabled',justify='center')
        cuadrota14.grid(row=2, column=0, pady=9,padx=10)

        def Agregar():
            global kkkinforme_tecnico 
            global kkkfecha_lectura 
            global kkkfecha_emision 
            global kkkcliente
            global kkkarea
            global kkkdias
            global kkkmes1 
            global kkkmes2 
            global kkkmes3 
            global kkkmes4 
            global kkkmes5 
            global kkkmes6 
            global kkkcargo_fijo_mensual 
            global kkkcargo_energia_activa_punta 
            global kkkcargo_energia_activa_fuera_punta 
            global kkkcargo_potencia_activa_generacion_presente_punta 
            global kkkcargo_potencia_activa_generacion_presente_fuera_punta 
            global kkkcargo_potencia_activa_redes_presente_punta 
            global kkkcargo_potencia_activa_redes_presente_fuera_punta 
            global kkkcargo_energia_reactiva_exc_30 
            global kkkcont_tableros
            global kkkfoto1
            global kkkfoto2
            global kkkfoto3
            global kkkmost_foto1
            global kkkmost_foto2
            global kkkmost_foto3
            global kkkfp_mes_actual 
            global kkkfp_mes_anterior 
            global kkkhp_mes_actual 
            global kkkhp_mes_anterior 
            global kkkea_mes_actual 
            global kkkea_mes_anterior 
            global kkkmaxima_demanda  
            global kkker_mes_actual 
            global kkker_mes_anterior
            global informe_tecnico 
            global tablero 
            global fecha_lectura 
            global fecha_emision
            global cliente 
            global area 
            global dias
            global meses 
            global promedio
            global cargo_fijo_mensual 
            global cargo_energia_activa_punta 
            global cargo_energia_activa_fuera_punta 
            global cargo_potencia_activa_generacion_presente_punta 
            global cargo_potencia_activa_generacion_presente_fuera_punta 
            global cargo_potencia_activa_redes_presente_punta 
            global cargo_potencia_activa_redes_presente_fuera_punta 
            global cargo_energia_reactiva_exc_30 
            global cont_tableros
            global foto1
            global foto2
            global foto3
            global fp_mes_actual 
            global fp_mes_anterior 
            global fp_mes 
            global hp_mes_actual  
            global hp_mes_anterior 
            global hp_mes 
            global ea_mes_actual  
            global ea_mes_anterior  
            global ea_mes 
            global maxima_demanda  
            global er_mes_actual  
            global er_mes_anterior  
            global er_mes 
            global fp_mestotal 
            global hp_mestotal 
            global ea_mestotal 
            global maxima_demandatotal 
            global maxima_demanda_del_mes
            global er_mestotal 
            global cali_tari
            global horas_punta 
            global texto
            global cargo_eapp 
            global cargo_eafp 
            global cargo_pagpp 
            global cargo_pagfp 
            global cargo_parpp 
            global cargo_parfp 
            global operacion1
            global operacion2 
            global operacion3 
            global operacion4 
            global v2
            global er_mestotalfac
            global texto2
            global v3
            global cargo_er_exc30 
            global v4
            global v5
            global v6
            global v7
            global subtotal
            global conigv
            global total_final
            global kkknom_tablero
            global nom_tablero

            if kkkfp_mes_actual.get()=="" or kkkfoto1=="" or kkkfoto2=="" or kkkfoto3=="" or kkkfp_mes_anterior.get()=="" or kkkhp_mes_actual.get()=="" or kkkhp_mes_anterior.get()=="" or kkkea_mes_actual.get()=="" or kkkea_mes_anterior.get()=="" or kkkmaxima_demanda.get()=="" or kkker_mes_actual.get()=="" or kkker_mes_anterior.get()=="" or kkknom_tablero.get()==""  :
                mensaje1()
                
            else:
                v1=round(float(kkkfp_mes_actual.get()),2)
                v2=round(float(kkkfp_mes_anterior.get()),2)
                v3=round(float(float(v1)-float(v2)),2)
                v4=round(float(kkkhp_mes_actual.get()),2)
                v5=round(float(kkkhp_mes_anterior.get()),2)
                v6=round(float(float(v4)-float(v5)),2)
                v7=round(float(float(v1)+float(v4)),2)
                v8=round(float(float(v2)+float(v5)),2)
                v9=round(float(float(v7)-float(v8)),2)
                v10=round(float(kkkmaxima_demanda.get()),2)
                v11=round(float(kkker_mes_actual.get()),2)
                v12=round(float(kkker_mes_anterior.get()),2)
                v13=round(float(float(v11)-float(v12)),2)
                ###################
                fp_mes_actual.append(float(v1))
                fp_mes_anterior.append(float(v2))
                fp_mes.append(round(float(v3),2))
                hp_mes_actual.append(float(v4))
                hp_mes_anterior.append(float(v5))
                hp_mes.append(round(float(v6),2))
                ea_mes_actual.append(float(v7))
                ea_mes_anterior.append(float(v8))
                ea_mes.append(round(float(v9),2))
                maxima_demanda.append(float(v10))
                er_mes_actual.append(float(v11))
                er_mes_anterior.append(float(v12))
                er_mes.append(round(float(v13),2))
                ##################
              
                foto1.append(str(kkkfoto1))
                foto2.append(str(kkkfoto2))
                foto3.append(str(kkkfoto3))

                nom_tablero.append(str(kkknom_tablero.get()))
                
                kkknom_tablero.set("")
                kkkfp_mes_actual.set("0")
                kkkfp_mes_anterior.set("0")
                kkkhp_mes_actual.set("0")
                kkkhp_mes_anterior.set("0")
                kkkea_mes_actual.set("0")
                kkkea_mes_anterior.set("0")
                kkkmaxima_demanda.set("0")
                kkker_mes_actual.set("0")
                kkker_mes_anterior.set("0")
                kkkmost_foto1.set("")
                kkkmost_foto2.set("")
                kkkmost_foto3.set("")

                cont_tableros=cont_tableros+1
                kkkcont_tableros.set(str(cont_tableros))
                if cont_tableros == 14:
                    btnagregar.configure(state='disabled')
                maxima_demanda_del_mes= float(v10)+maxima_demanda_del_mes
                kkkmes6.set(round(maxima_demanda_del_mes,2))


            
        btnagregar=Button(e111,text="Agregar", command= Agregar,font=("Arial", 10), relief="raised", borderwidth=4)
        btnagregar.grid(row=3, sticky="s",pady=7)

        #final
        framefinal= Frame(miframe,width=1200,height=50)
        framefinal.pack()
        framefinal.config(bg="#ecefee")


        estadolabel=Label(miframe, text="Estado: falta guardar        ",font=("Arial", 9)).place(x=10,y=760)
        estadolabel=Label(miframe, text="Derechos Reservados INSTCAL SAC       v1.0",font=("Arial", 8)).place(x=680,y=760)

        def asignar():
            global kkkinforme_tecnico 
            global kkkfecha_lectura 
            global kkkfecha_emision
            global kkkcliente
            global kkkarea
            global kkkdias
            global kkkmes1 
            global kkkmes2 
            global kkkmes3 
            global kkkmes4 
            global kkkmes5 
            global kkkmes6 
            global kkkcargo_fijo_mensual 
            global kkkcargo_energia_activa_punta 
            global kkkcargo_energia_activa_fuera_punta 
            global kkkcargo_potencia_activa_generacion_presente_punta 
            global kkkcargo_potencia_activa_generacion_presente_fuera_punta 
            global kkkcargo_potencia_activa_redes_presente_punta 
            global kkkcargo_potencia_activa_redes_presente_fuera_punta 
            global kkkcargo_energia_reactiva_exc_30 
            global kkkcont_tableros
            global kkkfoto1
            global kkkfoto2
            global kkkfoto3
            global kkkmost_foto1
            global kkkmost_foto2
            global kkkmost_foto3
            global kkkfp_mes_actual 
            global kkkfp_mes_anterior 
            global kkkhp_mes_actual 
            global kkkhp_mes_anterior 
            global kkkea_mes_actual 
            global kkkea_mes_anterior 
            global kkkmaxima_demanda  
            global kkker_mes_actual 
            global kkker_mes_anterior
            global informe_tecnico 
            global tablero 
            global fecha_lectura 
            global fecha_emision
            global cliente 
            global area 
            global dias
            global meses 
            global promedio
            global cargo_fijo_mensual 
            global cargo_energia_activa_punta 
            global cargo_energia_activa_fuera_punta 
            global cargo_potencia_activa_generacion_presente_punta 
            global cargo_potencia_activa_generacion_presente_fuera_punta 
            global cargo_potencia_activa_redes_presente_punta 
            global cargo_potencia_activa_redes_presente_fuera_punta 
            global cargo_energia_reactiva_exc_30 
            global cont_tableros
            global foto1
            global foto2
            global foto3
            global fp_mes_actual 
            global fp_mes_anterior 
            global fp_mes 
            global hp_mes_actual  
            global hp_mes_anterior 
            global hp_mes 
            global ea_mes_actual  
            global ea_mes_anterior  
            global ea_mes 
            global maxima_demanda  
            global er_mes_actual  
            global er_mes_anterior  
            global er_mes 
            global fp_mestotal 
            global hp_mestotal 
            global ea_mestotal 
            global maxima_demandatotal 
            global er_mestotal 
            global cali_tari
            global horas_punta 
            global texto
            global cargo_eapp 
            global cargo_eafp 
            global cargo_pagpp 
            global cargo_pagfp 
            global cargo_parpp 
            global cargo_parfp 
            global operacion1
            global operacion2 
            global operacion3 
            global operacion4 
            global v2
            global er_mestotalfac
            global texto2
            global v3
            global cargo_er_exc30 
            global v4
            global v5
            global v6
            global v7
            global subtotal
            global conigv
            global total_final
            global kkknom_tablero
            global nom_tablero
            global estadolabel
            global kkknom_mes1
            global kkknom_mes2
            global kkknom_mes3
            global kkknom_mes4
            global kkknom_mes5
            global kkknom_mes6
            global nom_mes1
            global nom_mes2
            global nom_mes3
            global nom_mes4
            global nom_mes5
            global nom_mes6
            global p1
            global p2
            global p3

            var_ordenar=[]

            if kkkdias.get()=="0" or kkkinforme_tecnico.get()=="" or kkkfecha_lectura.get()==""or kkkfecha_emision.get()==""or kkkfecha_lectura.get()=="dd-mm-yyyy"or kkkfecha_emision.get()=="dd-mm-yyyy" or kkkdias.get()=="" or kkkcliente.get()=="" or kkkarea.get()=="" or kkkmes1.get()=="" or kkkmes2.get()=="" or kkkmes3.get()=="" or kkkmes4.get()=="" or kkkmes5.get()=="" or kkkmes6.get()=="" or kkkcargo_fijo_mensual.get()=="" or kkkcargo_energia_activa_punta.get()=="" or kkkcargo_energia_activa_fuera_punta.get()=="" or kkkcargo_energia_reactiva_exc_30.get()=="" or kkkcargo_potencia_activa_generacion_presente_punta.get()=="" or kkkcargo_potencia_activa_generacion_presente_fuera_punta.get()=="" or kkkcargo_potencia_activa_redes_presente_punta.get()=="" or kkkcargo_potencia_activa_redes_presente_fuera_punta.get()=="" or cont_tableros<1 or kkknom_mes3.get()=="Mes" :
                mensaje2()
            else:
                separador= ', '
                tablero= separador.join(nom_tablero)
                tablero=tablero.upper()

                informe_tecnico=kkkinforme_tecnico.get()
                formato_fecha(kkkfecha_lectura.get(),1)
                formato_fecha(kkkfecha_emision.get(),2)
                cliente=kkkcliente.get()
                cliente=cliente.upper()
                area=kkkarea.get()
                area=area.upper()
                dias=int(kkkdias.get())

                nom_mes1=kkknom_mes1.get()
                nom_mes2=kkknom_mes2.get()
                nom_mes3=kkknom_mes3.get()
                nom_mes4=kkknom_mes4.get()
                nom_mes5=kkknom_mes5.get()
                nom_mes6=kkknom_mes6.get()

                meses.append(float(kkkmes1.get()))
                meses.append(float(kkkmes2.get()))
                meses.append(float(kkkmes3.get()))
                meses.append(float(kkkmes4.get()))
                meses.append(float(kkkmes5.get()))
                meses.append(float(kkkmes6.get()))

                var_ordenar.append(float(kkkmes1.get()))
                var_ordenar.append(float(kkkmes2.get()))
                var_ordenar.append(float(kkkmes3.get()))
                var_ordenar.append(float(kkkmes4.get()))
                var_ordenar.append(float(kkkmes5.get()))
                var_ordenar.append(float(kkkmes6.get()))
                var_ordenar.sort()

                if var_ordenar[4] == 0:
                        promedio=round(var_ordenar[5],2)
                else:
                        promedio = round(((var_ordenar[5]+var_ordenar[4])/2),2)
                cargo_fijo_mensual = round(float(kkkcargo_fijo_mensual.get()),2)
                cargo_energia_activa_punta = float(kkkcargo_energia_activa_punta.get())*0.01
                p1=round(float(kkkcargo_energia_activa_punta.get()),2)
                cargo_energia_activa_punta = round(cargo_energia_activa_punta,2)
                cargo_energia_activa_fuera_punta = float(kkkcargo_energia_activa_fuera_punta.get())*0.01
                p2=round(float(kkkcargo_energia_activa_fuera_punta.get()),2)
                cargo_energia_activa_fuera_punta = round(cargo_energia_activa_fuera_punta,2)
                cargo_potencia_activa_generacion_presente_punta = float(kkkcargo_potencia_activa_generacion_presente_punta.get())
                cargo_potencia_activa_generacion_presente_fuera_punta = round(float(kkkcargo_potencia_activa_generacion_presente_fuera_punta.get()),2)
                cargo_potencia_activa_redes_presente_punta = float(kkkcargo_potencia_activa_redes_presente_punta.get())
                cargo_potencia_activa_redes_presente_fuera_punta = round(float(kkkcargo_potencia_activa_redes_presente_fuera_punta.get()),2)
                cargo_energia_reactiva_exc_30 = float(kkkcargo_energia_reactiva_exc_30.get())
                p3=round(float(kkkcargo_energia_reactiva_exc_30.get()),2)
                cargo_energia_reactiva_exc_30 = round(cargo_energia_reactiva_exc_30,2)

                # OBTENIENDO SUMA DE LAS VARIABLES DE TODOS LOS MEDIDORES 
                fp_mestotal= 0
                hp_mestotal= 0
                ea_mestotal= 0
                maxima_demandatotal= 0
                er_mestotal= 0
                for i in range(cont_tableros):

                    fp_mestotal= (fp_mestotal + fp_mes[i])

                    hp_mestotal= (hp_mestotal + hp_mes[i])

                    ea_mestotal= (ea_mestotal + ea_mes[i])
                    
                    maxima_demandatotal= (maxima_demandatotal + maxima_demanda[i])

                    er_mestotal= (er_mestotal + er_mes[i])

                fp_mestotal= round(fp_mestotal,2)

                hp_mestotal= round(hp_mestotal,2)

                ea_mestotal= round(ea_mestotal,2)
                    
                maxima_demandatotal= round(maxima_demandatotal,2)

                er_mestotal= round(er_mestotal,2)
                
                horas_punta = (5*dias)

                if maxima_demandatotal == 0 or horas_punta==0:
                    cali_tari=0.0
                    print("horas puntaaaa", horas_punta)
                    print("max", maxima_demandatotal)

                else:
                    cali_tari= round(float(hp_mestotal/(maxima_demandatotal*horas_punta )),2)
                    print("horas punta", horas_punta)
                    print("max", maxima_demandatotal)
                    print(cali_tari)

                
                if cali_tari >= 0.50:
                        texto= "CLIENTE PRESENTE EN PUNTA"
                        v=1
                else:
                        texto= "CLIENTE FUERA DE PUNTA"
                        v=0

                cargo_eapp = round(float((hp_mestotal * p1 * 0.01)),2)
                cargo_eafp = round(float((fp_mestotal * p2 * 0.01)),2)
                cargo_pagpp = round(float((maxima_demandatotal * cargo_potencia_activa_generacion_presente_punta)),2)
                cargo_pagfp = round(float((maxima_demandatotal * cargo_potencia_activa_generacion_presente_fuera_punta)),2)
                cargo_parpp = round(float((promedio * cargo_potencia_activa_redes_presente_punta)),2)
                cargo_parfp = round(float((promedio * cargo_potencia_activa_redes_presente_fuera_punta)),2)

                operacion1 =str(maxima_demandatotal)+' Kw  x '+str(cargo_potencia_activa_generacion_presente_punta)+' S/./kW.h = '+str(cargo_pagpp)+' S/./kW.h'
                operacion2 =str(maxima_demandatotal)+' Kw  x '+str(cargo_potencia_activa_generacion_presente_fuera_punta)+' S/./kW.h = '+str(cargo_pagfp)+' S/./kW.h'
                operacion3 =str(promedio)+' Kw  x '+str(cargo_potencia_activa_redes_presente_punta)+' S/./kW.h = '+str(cargo_parpp)+' S/./kW.h'
                operacion4 =str(promedio)+' Kw  x '+str(cargo_potencia_activa_redes_presente_fuera_punta)+' S/./kW.h = '+str(cargo_parfp)+' S/./kW.h'

                if v == 0:
                        operacion1= ""
                        operacion3= ""
                        cargo_pagpp=0
                        cargo_parpp=0
                else:
                        operacion2= ""
                        operacion4= ""
                        cargo_pagfp=0
                        cargo_parfp=0

                v2=round((0.3*ea_mestotal),2)
                er_mestotalfac=round((er_mestotal-(0.3*ea_mestotal)),2)

                if er_mestotalfac > 0:
                        texto2= "Se factura cargos por energia reactiva, se cumple la condición inicial."
                        v3=round(er_mestotalfac,2)

                else:
                        texto2= "No se factura cargos por energia reactiva, no se cumple la condición inicial."
                        v3=round(er_mestotalfac,2)
                        er_mestotalfac=0

                cargo_er_exc30 =round((er_mestotalfac * cargo_energia_reactiva_exc_30 ),2)

                if v == 0:
                        v4= 0
                        v5= maxima_demandatotal
                        v6= 0
                        v7= promedio
                else:
                        v4= maxima_demandatotal
                        v5= 0
                        v6= promedio
                        v7= 0

                subtotal= round((cargo_fijo_mensual+cargo_eapp+cargo_eafp+cargo_pagpp+cargo_pagfp+cargo_parpp+cargo_parfp+cargo_er_exc30),2)
                conigv= round((subtotal*0.18),2)
                total_final= round((subtotal+conigv),2)

                estadolabel=Label(miframe, text="Estado: Guardado exitoso",font=("Arial", 10)).place(x=10,y=760)
                btnfinal2.configure(state="normal")


        def Exportar():
            global kkkinforme_tecnico 
            global kkkfecha_lectura 
            global kkkfecha_emision
            global kkkcliente
            global kkkarea
            global kkkdias
            global kkkmes1 
            global kkkmes2 
            global kkkmes3 
            global kkkmes4 
            global kkkmes5 
            global kkkmes6 
            global kkkcargo_fijo_mensual 
            global kkkcargo_energia_activa_punta 
            global kkkcargo_energia_activa_fuera_punta 
            global kkkcargo_potencia_activa_generacion_presente_punta 
            global kkkcargo_potencia_activa_generacion_presente_fuera_punta 
            global kkkcargo_potencia_activa_redes_presente_punta 
            global kkkcargo_potencia_activa_redes_presente_fuera_punta 
            global kkkcargo_energia_reactiva_exc_30 
            global kkkcont_tableros
            global kkkfoto1
            global kkkfoto2
            global kkkfoto3
            global kkkmost_foto1
            global kkkmost_foto2
            global kkkmost_foto3
            global kkkfp_mes_actual 
            global kkkfp_mes_anterior 
            global kkkhp_mes_actual 
            global kkkhp_mes_anterior 
            global kkkea_mes_actual 
            global kkkea_mes_anterior 
            global kkkmaxima_demanda  
            global kkker_mes_actual 
            global kkker_mes_anterior
            global informe_tecnico 
            global tablero 
            global fecha_lectura 
            global fecha_emision
            global cliente 
            global area 
            global dias
            global meses 
            global promedio
            global cargo_fijo_mensual 
            global cargo_energia_activa_punta 
            global cargo_energia_activa_fuera_punta 
            global cargo_potencia_activa_generacion_presente_punta 
            global cargo_potencia_activa_generacion_presente_fuera_punta 
            global cargo_potencia_activa_redes_presente_punta 
            global cargo_potencia_activa_redes_presente_fuera_punta 
            global cargo_energia_reactiva_exc_30 
            global cont_tableros
            global foto1
            global foto2
            global foto3
            global fp_mes_actual 
            global fp_mes_anterior 
            global fp_mes 
            global hp_mes_actual  
            global hp_mes_anterior 
            global hp_mes 
            global ea_mes_actual  
            global ea_mes_anterior  
            global ea_mes 
            global maxima_demanda  
            global er_mes_actual  
            global er_mes_anterior  
            global er_mes 
            global fp_mestotal 
            global hp_mestotal 
            global ea_mestotal 
            global maxima_demandatotal 
            global er_mestotal 
            global cali_tari
            global horas_punta 
            global texto
            global cargo_eapp 
            global cargo_eafp 
            global cargo_pagpp 
            global cargo_pagfp 
            global cargo_parpp 
            global cargo_parfp 
            global operacion1
            global operacion2 
            global operacion3 
            global operacion4 
            global v2
            global er_mestotalfac
            global texto2
            global v3
            global cargo_er_exc30 
            global v4
            global v5
            global v6
            global v7
            global subtotal
            global conigv
            global total_final
            global kkknom_tablero
            global nom_tablero
            global nom_mes1
            global nom_mes2
            global nom_mes3
            global nom_mes4
            global nom_mes5
            global nom_mes6
            global anio
            global p1
            global p2
            global p3

            if cont_tableros == 1:
                doc = DocxTemplate("templates/template1.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,

                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2) ,        

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 1")
            if cont_tableros == 2:
                doc = DocxTemplate("templates/template2.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,

                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 2")
            if cont_tableros == 3:
                doc = DocxTemplate("templates/template3.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),
                
                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 3")
            if cont_tableros == 4:
                doc = DocxTemplate("templates/template4.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 4")
            if cont_tableros == 5:
                doc = DocxTemplate("templates/template5.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),
                
                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 5")
            if cont_tableros == 6:
                doc = DocxTemplate("templates/template6.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act5": fp_mes_actual[5] ,
                "fp_mes_ant5": fp_mes_anterior[5] ,
                "fp_mes5": fp_mes[5],
                "hp_mes_act5": hp_mes_actual[5] ,
                "hp_mes_ant5": hp_mes_anterior[5] ,
                "hp_mes5": hp_mes[5] ,
                "ea_mes_act5": ea_mes_actual[5] ,
                "ea_mes_ant5": ea_mes_anterior[5] ,
                "ea_mes5": ea_mes[5] ,
                "maxima_demanda5": maxima_demanda[5] ,
                "er_mes_act5": er_mes_actual[5] ,
                "er_mes_ant5": er_mes_anterior[5] ,
                "er_mes5": er_mes[5] ,
                "nom_tablero5":nom_tablero[5] ,
                "imgab5": InlineImage(doc,foto1[5],height=Mm(35), width=Mm(45)) ,
                "imgc5": InlineImage(doc,foto2[5],height=Mm(35), width=Mm(45)) ,
                "imgd5": InlineImage(doc,foto3[5],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 6")
            if cont_tableros == 7:
                doc = DocxTemplate("templates/template7.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act5": fp_mes_actual[5] ,
                "fp_mes_ant5": fp_mes_anterior[5] ,
                "fp_mes5": fp_mes[5],
                "hp_mes_act5": hp_mes_actual[5] ,
                "hp_mes_ant5": hp_mes_anterior[5] ,
                "hp_mes5": hp_mes[5] ,
                "ea_mes_act5": ea_mes_actual[5] ,
                "ea_mes_ant5": ea_mes_anterior[5] ,
                "ea_mes5": ea_mes[5] ,
                "maxima_demanda5": maxima_demanda[5] ,
                "er_mes_act5": er_mes_actual[5] ,
                "er_mes_ant5": er_mes_anterior[5] ,
                "er_mes5": er_mes[5] ,
                "nom_tablero5":nom_tablero[5] ,
                "imgab5": InlineImage(doc,foto1[5],height=Mm(35), width=Mm(45)) ,
                "imgc5": InlineImage(doc,foto2[5],height=Mm(35), width=Mm(45)) ,
                "imgd5": InlineImage(doc,foto3[5],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act6": fp_mes_actual[6] ,
                "fp_mes_ant6": fp_mes_anterior[6] ,
                "fp_mes6": fp_mes[6] ,
                "hp_mes_act6": hp_mes_actual[6] ,
                "hp_mes_ant6": hp_mes_anterior[6] ,
                "hp_mes6": hp_mes[6] ,
                "ea_mes_act6": ea_mes_actual[6] ,
                "ea_mes_ant6": ea_mes_anterior[6] ,
                "ea_mes6": ea_mes[6] ,
                "maxima_demanda6": maxima_demanda[6] ,
                "er_mes_act6": er_mes_actual[6] ,
                "er_mes_ant6": er_mes_anterior[6] ,
                "er_mes6": er_mes[6] ,
                "nom_tablero6":nom_tablero[6] ,
                "imgab6": InlineImage(doc,foto1[6],height=Mm(35), width=Mm(45)) ,
                "imgc6": InlineImage(doc,foto2[6],height=Mm(35), width=Mm(45)) ,
                "imgd6": InlineImage(doc,foto3[6],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 7")
            if cont_tableros == 8:
                doc = DocxTemplate("templates/template8.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act5": fp_mes_actual[5] ,
                "fp_mes_ant5": fp_mes_anterior[5] ,
                "fp_mes5": fp_mes[5],
                "hp_mes_act5": hp_mes_actual[5] ,
                "hp_mes_ant5": hp_mes_anterior[5] ,
                "hp_mes5": hp_mes[5] ,
                "ea_mes_act5": ea_mes_actual[5] ,
                "ea_mes_ant5": ea_mes_anterior[5] ,
                "ea_mes5": ea_mes[5] ,
                "maxima_demanda5": maxima_demanda[5] ,
                "er_mes_act5": er_mes_actual[5] ,
                "er_mes_ant5": er_mes_anterior[5] ,
                "er_mes5": er_mes[5] ,
                "nom_tablero5":nom_tablero[5] ,
                "imgab5": InlineImage(doc,foto1[5],height=Mm(35), width=Mm(45)) ,
                "imgc5": InlineImage(doc,foto2[5],height=Mm(35), width=Mm(45)) ,
                "imgd5": InlineImage(doc,foto3[5],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act6": fp_mes_actual[6] ,
                "fp_mes_ant6": fp_mes_anterior[6] ,
                "fp_mes6": fp_mes[6] ,
                "hp_mes_act6": hp_mes_actual[6] ,
                "hp_mes_ant6": hp_mes_anterior[6] ,
                "hp_mes6": hp_mes[6] ,
                "ea_mes_act6": ea_mes_actual[6] ,
                "ea_mes_ant6": ea_mes_anterior[6] ,
                "ea_mes6": ea_mes[6] ,
                "maxima_demanda6": maxima_demanda[6] ,
                "er_mes_act6": er_mes_actual[6] ,
                "er_mes_ant6": er_mes_anterior[6] ,
                "er_mes6": er_mes[6] ,
                "nom_tablero6":nom_tablero[6] ,
                "imgab6": InlineImage(doc,foto1[6],height=Mm(35), width=Mm(45)) ,
                "imgc6": InlineImage(doc,foto2[6],height=Mm(35), width=Mm(45)) ,
                "imgd6": InlineImage(doc,foto3[6],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act7": fp_mes_actual[7] ,
                "fp_mes_ant7": fp_mes_anterior[7] ,
                "fp_mes7": fp_mes[7] ,
                "hp_mes_act7": hp_mes_actual[7] ,
                "hp_mes_ant7": hp_mes_anterior[7] ,
                "hp_mes7": hp_mes[7] ,
                "ea_mes_act7": ea_mes_actual[7] ,
                "ea_mes_ant7": ea_mes_anterior[7] ,
                "ea_mes7": ea_mes[7] ,
                "maxima_demanda7": maxima_demanda[7] ,
                "er_mes_act7": er_mes_actual[7] ,
                "er_mes_ant7": er_mes_anterior[7] ,
                "er_mes7": er_mes[7] ,
                "nom_tablero7":nom_tablero[7] ,
                "imgab7": InlineImage(doc,foto1[7],height=Mm(35), width=Mm(45)) ,
                "imgc7": InlineImage(doc,foto2[7],height=Mm(35), width=Mm(45)) ,
                "imgd7": InlineImage(doc,foto3[7],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 8")
            if cont_tableros == 9:
                doc = DocxTemplate("templates/template9.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act5": fp_mes_actual[5] ,
                "fp_mes_ant5": fp_mes_anterior[5] ,
                "fp_mes5": fp_mes[5],
                "hp_mes_act5": hp_mes_actual[5] ,
                "hp_mes_ant5": hp_mes_anterior[5] ,
                "hp_mes5": hp_mes[5] ,
                "ea_mes_act5": ea_mes_actual[5] ,
                "ea_mes_ant5": ea_mes_anterior[5] ,
                "ea_mes5": ea_mes[5] ,
                "maxima_demanda5": maxima_demanda[5] ,
                "er_mes_act5": er_mes_actual[5] ,
                "er_mes_ant5": er_mes_anterior[5] ,
                "er_mes5": er_mes[5] ,
                "nom_tablero5":nom_tablero[5] ,
                "imgab5": InlineImage(doc,foto1[5],height=Mm(35), width=Mm(45)) ,
                "imgc5": InlineImage(doc,foto2[5],height=Mm(35), width=Mm(45)) ,
                "imgd5": InlineImage(doc,foto3[5],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act6": fp_mes_actual[6] ,
                "fp_mes_ant6": fp_mes_anterior[6] ,
                "fp_mes6": fp_mes[6] ,
                "hp_mes_act6": hp_mes_actual[6] ,
                "hp_mes_ant6": hp_mes_anterior[6] ,
                "hp_mes6": hp_mes[6] ,
                "ea_mes_act6": ea_mes_actual[6] ,
                "ea_mes_ant6": ea_mes_anterior[6] ,
                "ea_mes6": ea_mes[6] ,
                "maxima_demanda6": maxima_demanda[6] ,
                "er_mes_act6": er_mes_actual[6] ,
                "er_mes_ant6": er_mes_anterior[6] ,
                "er_mes6": er_mes[6] ,
                "nom_tablero6":nom_tablero[6] ,
                "imgab6": InlineImage(doc,foto1[6],height=Mm(35), width=Mm(45)) ,
                "imgc6": InlineImage(doc,foto2[6],height=Mm(35), width=Mm(45)) ,
                "imgd6": InlineImage(doc,foto3[6],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act7": fp_mes_actual[7] ,
                "fp_mes_ant7": fp_mes_anterior[7] ,
                "fp_mes7": fp_mes[7] ,
                "hp_mes_act7": hp_mes_actual[7] ,
                "hp_mes_ant7": hp_mes_anterior[7] ,
                "hp_mes7": hp_mes[7] ,
                "ea_mes_act7": ea_mes_actual[7] ,
                "ea_mes_ant7": ea_mes_anterior[7] ,
                "ea_mes7": ea_mes[7] ,
                "maxima_demanda7": maxima_demanda[7] ,
                "er_mes_act7": er_mes_actual[7] ,
                "er_mes_ant7": er_mes_anterior[7] ,
                "er_mes7": er_mes[7] ,
                "nom_tablero7":nom_tablero[7] ,
                "imgab7": InlineImage(doc,foto1[7],height=Mm(35), width=Mm(45)) ,
                "imgc7": InlineImage(doc,foto2[7],height=Mm(35), width=Mm(45)) ,
                "imgd7": InlineImage(doc,foto3[7],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act8": fp_mes_actual[8] ,
                "fp_mes_ant8": fp_mes_anterior[8] ,
                "fp_mes8": fp_mes[8] ,
                "hp_mes_act8": hp_mes_actual[8] ,
                "hp_mes_ant8": hp_mes_anterior[8] ,
                "hp_mes8": hp_mes[8] ,
                "ea_mes_act8": ea_mes_actual[8] ,
                "ea_mes_ant8": ea_mes_anterior[8] ,
                "ea_mes8": ea_mes[8] ,
                "maxima_demanda8": maxima_demanda[8] ,
                "er_mes_act8": er_mes_actual[8] ,
                "er_mes_ant8": er_mes_anterior[8] ,
                "er_mes8": er_mes[8] ,
                "nom_tablero8":nom_tablero[8] ,
                "imgab8": InlineImage(doc,foto1[8],height=Mm(35), width=Mm(45)) ,
                "imgc8": InlineImage(doc,foto2[8],height=Mm(35), width=Mm(45)) ,
                "imgd8": InlineImage(doc,foto3[8],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 9")
            if cont_tableros == 10:
                doc = DocxTemplate("templates/template10.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act5": fp_mes_actual[5] ,
                "fp_mes_ant5": fp_mes_anterior[5] ,
                "fp_mes5": fp_mes[5],
                "hp_mes_act5": hp_mes_actual[5] ,
                "hp_mes_ant5": hp_mes_anterior[5] ,
                "hp_mes5": hp_mes[5] ,
                "ea_mes_act5": ea_mes_actual[5] ,
                "ea_mes_ant5": ea_mes_anterior[5] ,
                "ea_mes5": ea_mes[5] ,
                "maxima_demanda5": maxima_demanda[5] ,
                "er_mes_act5": er_mes_actual[5] ,
                "er_mes_ant5": er_mes_anterior[5] ,
                "er_mes5": er_mes[5] ,
                "nom_tablero5":nom_tablero[5] ,
                "imgab5": InlineImage(doc,foto1[5],height=Mm(35), width=Mm(45)) ,
                "imgc5": InlineImage(doc,foto2[5],height=Mm(35), width=Mm(45)) ,
                "imgd5": InlineImage(doc,foto3[5],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act6": fp_mes_actual[6] ,
                "fp_mes_ant6": fp_mes_anterior[6] ,
                "fp_mes6": fp_mes[6] ,
                "hp_mes_act6": hp_mes_actual[6] ,
                "hp_mes_ant6": hp_mes_anterior[6] ,
                "hp_mes6": hp_mes[6] ,
                "ea_mes_act6": ea_mes_actual[6] ,
                "ea_mes_ant6": ea_mes_anterior[6] ,
                "ea_mes6": ea_mes[6] ,
                "maxima_demanda6": maxima_demanda[6] ,
                "er_mes_act6": er_mes_actual[6] ,
                "er_mes_ant6": er_mes_anterior[6] ,
                "er_mes6": er_mes[6] ,
                "nom_tablero6":nom_tablero[6] ,
                "imgab6": InlineImage(doc,foto1[6],height=Mm(35), width=Mm(45)) ,
                "imgc6": InlineImage(doc,foto2[6],height=Mm(35), width=Mm(45)) ,
                "imgd6": InlineImage(doc,foto3[6],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act7": fp_mes_actual[7] ,
                "fp_mes_ant7": fp_mes_anterior[7] ,
                "fp_mes7": fp_mes[7] ,
                "hp_mes_act7": hp_mes_actual[7] ,
                "hp_mes_ant7": hp_mes_anterior[7] ,
                "hp_mes7": hp_mes[7] ,
                "ea_mes_act7": ea_mes_actual[7] ,
                "ea_mes_ant7": ea_mes_anterior[7] ,
                "ea_mes7": ea_mes[7] ,
                "maxima_demanda7": maxima_demanda[7] ,
                "er_mes_act7": er_mes_actual[7] ,
                "er_mes_ant7": er_mes_anterior[7] ,
                "er_mes7": er_mes[7] ,
                "nom_tablero7":nom_tablero[7] ,
                "imgab7": InlineImage(doc,foto1[7],height=Mm(35), width=Mm(45)) ,
                "imgc7": InlineImage(doc,foto2[7],height=Mm(35), width=Mm(45)) ,
                "imgd7": InlineImage(doc,foto3[7],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act8": fp_mes_actual[8] ,
                "fp_mes_ant8": fp_mes_anterior[8] ,
                "fp_mes8": fp_mes[8] ,
                "hp_mes_act8": hp_mes_actual[8] ,
                "hp_mes_ant8": hp_mes_anterior[8] ,
                "hp_mes8": hp_mes[8] ,
                "ea_mes_act8": ea_mes_actual[8] ,
                "ea_mes_ant8": ea_mes_anterior[8] ,
                "ea_mes8": ea_mes[8] ,
                "maxima_demanda8": maxima_demanda[8] ,
                "er_mes_act8": er_mes_actual[8] ,
                "er_mes_ant8": er_mes_anterior[8] ,
                "er_mes8": er_mes[8] ,
                "nom_tablero8":nom_tablero[8] ,
                "imgab8": InlineImage(doc,foto1[8],height=Mm(35), width=Mm(45)) ,
                "imgc8": InlineImage(doc,foto2[8],height=Mm(35), width=Mm(45)) ,
                "imgd8": InlineImage(doc,foto3[8],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act9": fp_mes_actual[9] ,
                "fp_mes_ant9": fp_mes_anterior[9] ,
                "fp_mes9": fp_mes[9] ,
                "hp_mes_act9": hp_mes_actual[9] ,
                "hp_mes_ant9": hp_mes_anterior[9] ,
                "hp_mes9": hp_mes[9] ,
                "ea_mes_act9": ea_mes_actual[9] ,
                "ea_mes_ant9": ea_mes_anterior[9] ,
                "ea_mes9": ea_mes[9] ,
                "maxima_demanda9": maxima_demanda[9] ,
                "er_mes_act9": er_mes_actual[9] ,
                "er_mes_ant9": er_mes_anterior[9] ,
                "er_mes9": er_mes[9] ,
                "nom_tablero9":nom_tablero[9] ,
                "imgab9": InlineImage(doc,foto1[9],height=Mm(35), width=Mm(45)) ,
                "imgc9": InlineImage(doc,foto2[9],height=Mm(35), width=Mm(45)) ,
                "imgd9": InlineImage(doc,foto3[9],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 10")
            if cont_tableros == 11:
                doc = DocxTemplate("templates/template11.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act5": fp_mes_actual[5] ,
                "fp_mes_ant5": fp_mes_anterior[5] ,
                "fp_mes5": fp_mes[5],
                "hp_mes_act5": hp_mes_actual[5] ,
                "hp_mes_ant5": hp_mes_anterior[5] ,
                "hp_mes5": hp_mes[5] ,
                "ea_mes_act5": ea_mes_actual[5] ,
                "ea_mes_ant5": ea_mes_anterior[5] ,
                "ea_mes5": ea_mes[5] ,
                "maxima_demanda5": maxima_demanda[5] ,
                "er_mes_act5": er_mes_actual[5] ,
                "er_mes_ant5": er_mes_anterior[5] ,
                "er_mes5": er_mes[5] ,
                "nom_tablero5":nom_tablero[5] ,
                "imgab5": InlineImage(doc,foto1[5],height=Mm(35), width=Mm(45)) ,
                "imgc5": InlineImage(doc,foto2[5],height=Mm(35), width=Mm(45)) ,
                "imgd5": InlineImage(doc,foto3[5],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act6": fp_mes_actual[6] ,
                "fp_mes_ant6": fp_mes_anterior[6] ,
                "fp_mes6": fp_mes[6] ,
                "hp_mes_act6": hp_mes_actual[6] ,
                "hp_mes_ant6": hp_mes_anterior[6] ,
                "hp_mes6": hp_mes[6] ,
                "ea_mes_act6": ea_mes_actual[6] ,
                "ea_mes_ant6": ea_mes_anterior[6] ,
                "ea_mes6": ea_mes[6] ,
                "maxima_demanda6": maxima_demanda[6] ,
                "er_mes_act6": er_mes_actual[6] ,
                "er_mes_ant6": er_mes_anterior[6] ,
                "er_mes6": er_mes[6] ,
                "nom_tablero6":nom_tablero[6] ,
                "imgab6": InlineImage(doc,foto1[6],height=Mm(35), width=Mm(45)) ,
                "imgc6": InlineImage(doc,foto2[6],height=Mm(35), width=Mm(45)) ,
                "imgd6": InlineImage(doc,foto3[6],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act7": fp_mes_actual[7] ,
                "fp_mes_ant7": fp_mes_anterior[7] ,
                "fp_mes7": fp_mes[7] ,
                "hp_mes_act7": hp_mes_actual[7] ,
                "hp_mes_ant7": hp_mes_anterior[7] ,
                "hp_mes7": hp_mes[7] ,
                "ea_mes_act7": ea_mes_actual[7] ,
                "ea_mes_ant7": ea_mes_anterior[7] ,
                "ea_mes7": ea_mes[7] ,
                "maxima_demanda7": maxima_demanda[7] ,
                "er_mes_act7": er_mes_actual[7] ,
                "er_mes_ant7": er_mes_anterior[7] ,
                "er_mes7": er_mes[7] ,
                "nom_tablero7":nom_tablero[7] ,
                "imgab7": InlineImage(doc,foto1[7],height=Mm(35), width=Mm(45)) ,
                "imgc7": InlineImage(doc,foto2[7],height=Mm(35), width=Mm(45)) ,
                "imgd7": InlineImage(doc,foto3[7],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act8": fp_mes_actual[8] ,
                "fp_mes_ant8": fp_mes_anterior[8] ,
                "fp_mes8": fp_mes[8] ,
                "hp_mes_act8": hp_mes_actual[8] ,
                "hp_mes_ant8": hp_mes_anterior[8] ,
                "hp_mes8": hp_mes[8] ,
                "ea_mes_act8": ea_mes_actual[8] ,
                "ea_mes_ant8": ea_mes_anterior[8] ,
                "ea_mes8": ea_mes[8] ,
                "maxima_demanda8": maxima_demanda[8] ,
                "er_mes_act8": er_mes_actual[8] ,
                "er_mes_ant8": er_mes_anterior[8] ,
                "er_mes8": er_mes[8] ,
                "nom_tablero8":nom_tablero[8] ,
                "imgab8": InlineImage(doc,foto1[8],height=Mm(35), width=Mm(45)) ,
                "imgc8": InlineImage(doc,foto2[8],height=Mm(35), width=Mm(45)) ,
                "imgd8": InlineImage(doc,foto3[8],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act9": fp_mes_actual[9] ,
                "fp_mes_ant9": fp_mes_anterior[9] ,
                "fp_mes9": fp_mes[9] ,
                "hp_mes_act9": hp_mes_actual[9] ,
                "hp_mes_ant9": hp_mes_anterior[9] ,
                "hp_mes9": hp_mes[9] ,
                "ea_mes_act9": ea_mes_actual[9] ,
                "ea_mes_ant9": ea_mes_anterior[9] ,
                "ea_mes9": ea_mes[9] ,
                "maxima_demanda9": maxima_demanda[9] ,
                "er_mes_act9": er_mes_actual[9] ,
                "er_mes_ant9": er_mes_anterior[9] ,
                "er_mes9": er_mes[9] ,
                "nom_tablero9":nom_tablero[9] ,
                "imgab9": InlineImage(doc,foto1[9],height=Mm(35), width=Mm(45)) ,
                "imgc9": InlineImage(doc,foto2[9],height=Mm(35), width=Mm(45)) ,
                "imgd9": InlineImage(doc,foto3[9],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act10": fp_mes_actual[10] ,
                "fp_mes_ant10": fp_mes_anterior[10] ,
                "fp_mes10": fp_mes[10] ,
                "hp_mes_act10": hp_mes_actual[10] ,
                "hp_mes_ant10": hp_mes_anterior[10] ,
                "hp_mes10": hp_mes[10] ,
                "ea_mes_act10": ea_mes_actual[10] ,
                "ea_mes_ant10": ea_mes_anterior[10] ,
                "ea_mes10": ea_mes[10] ,
                "maxima_demanda10": maxima_demanda[10] ,
                "er_mes_act10": er_mes_actual[10] ,
                "er_mes_ant10": er_mes_anterior[10] ,
                "er_mes10": er_mes[10] ,
                "nom_tablero10":nom_tablero[10] ,
                "imgab10": InlineImage(doc,foto1[10],height=Mm(35), width=Mm(45)) ,
                "imgc10": InlineImage(doc,foto2[10],height=Mm(35), width=Mm(45)) ,
                "imgd10": InlineImage(doc,foto3[10],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 11")
            if cont_tableros == 12:
                doc = DocxTemplate("templates/template12.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act5": fp_mes_actual[5] ,
                "fp_mes_ant5": fp_mes_anterior[5] ,
                "fp_mes5": fp_mes[5],
                "hp_mes_act5": hp_mes_actual[5] ,
                "hp_mes_ant5": hp_mes_anterior[5] ,
                "hp_mes5": hp_mes[5] ,
                "ea_mes_act5": ea_mes_actual[5] ,
                "ea_mes_ant5": ea_mes_anterior[5] ,
                "ea_mes5": ea_mes[5] ,
                "maxima_demanda5": maxima_demanda[5] ,
                "er_mes_act5": er_mes_actual[5] ,
                "er_mes_ant5": er_mes_anterior[5] ,
                "er_mes5": er_mes[5] ,
                "nom_tablero5":nom_tablero[5] ,
                "imgab5": InlineImage(doc,foto1[5],height=Mm(35), width=Mm(45)) ,
                "imgc5": InlineImage(doc,foto2[5],height=Mm(35), width=Mm(45)) ,
                "imgd5": InlineImage(doc,foto3[5],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act6": fp_mes_actual[6] ,
                "fp_mes_ant6": fp_mes_anterior[6] ,
                "fp_mes6": fp_mes[6] ,
                "hp_mes_act6": hp_mes_actual[6] ,
                "hp_mes_ant6": hp_mes_anterior[6] ,
                "hp_mes6": hp_mes[6] ,
                "ea_mes_act6": ea_mes_actual[6] ,
                "ea_mes_ant6": ea_mes_anterior[6] ,
                "ea_mes6": ea_mes[6] ,
                "maxima_demanda6": maxima_demanda[6] ,
                "er_mes_act6": er_mes_actual[6] ,
                "er_mes_ant6": er_mes_anterior[6] ,
                "er_mes6": er_mes[6] ,
                "nom_tablero6":nom_tablero[6] ,
                "imgab6": InlineImage(doc,foto1[6],height=Mm(35), width=Mm(45)) ,
                "imgc6": InlineImage(doc,foto2[6],height=Mm(35), width=Mm(45)) ,
                "imgd6": InlineImage(doc,foto3[6],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act7": fp_mes_actual[7] ,
                "fp_mes_ant7": fp_mes_anterior[7] ,
                "fp_mes7": fp_mes[7] ,
                "hp_mes_act7": hp_mes_actual[7] ,
                "hp_mes_ant7": hp_mes_anterior[7] ,
                "hp_mes7": hp_mes[7] ,
                "ea_mes_act7": ea_mes_actual[7] ,
                "ea_mes_ant7": ea_mes_anterior[7] ,
                "ea_mes7": ea_mes[7] ,
                "maxima_demanda7": maxima_demanda[7] ,
                "er_mes_act7": er_mes_actual[7] ,
                "er_mes_ant7": er_mes_anterior[7] ,
                "er_mes7": er_mes[7] ,
                "nom_tablero7":nom_tablero[7] ,
                "imgab7": InlineImage(doc,foto1[7],height=Mm(35), width=Mm(45)) ,
                "imgc7": InlineImage(doc,foto2[7],height=Mm(35), width=Mm(45)) ,
                "imgd7": InlineImage(doc,foto3[7],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act8": fp_mes_actual[8] ,
                "fp_mes_ant8": fp_mes_anterior[8] ,
                "fp_mes8": fp_mes[8] ,
                "hp_mes_act8": hp_mes_actual[8] ,
                "hp_mes_ant8": hp_mes_anterior[8] ,
                "hp_mes8": hp_mes[8] ,
                "ea_mes_act8": ea_mes_actual[8] ,
                "ea_mes_ant8": ea_mes_anterior[8] ,
                "ea_mes8": ea_mes[8] ,
                "maxima_demanda8": maxima_demanda[8] ,
                "er_mes_act8": er_mes_actual[8] ,
                "er_mes_ant8": er_mes_anterior[8] ,
                "er_mes8": er_mes[8] ,
                "nom_tablero8":nom_tablero[8] ,
                "imgab8": InlineImage(doc,foto1[8],height=Mm(35), width=Mm(45)) ,
                "imgc8": InlineImage(doc,foto2[8],height=Mm(35), width=Mm(45)) ,
                "imgd8": InlineImage(doc,foto3[8],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act9": fp_mes_actual[9] ,
                "fp_mes_ant9": fp_mes_anterior[9] ,
                "fp_mes9": fp_mes[9] ,
                "hp_mes_act9": hp_mes_actual[9] ,
                "hp_mes_ant9": hp_mes_anterior[9] ,
                "hp_mes9": hp_mes[9] ,
                "ea_mes_act9": ea_mes_actual[9] ,
                "ea_mes_ant9": ea_mes_anterior[9] ,
                "ea_mes9": ea_mes[9] ,
                "maxima_demanda9": maxima_demanda[9] ,
                "er_mes_act9": er_mes_actual[9] ,
                "er_mes_ant9": er_mes_anterior[9] ,
                "er_mes9": er_mes[9] ,
                "nom_tablero9":nom_tablero[9] ,
                "imgab9": InlineImage(doc,foto1[9],height=Mm(35), width=Mm(45)) ,
                "imgc9": InlineImage(doc,foto2[9],height=Mm(35), width=Mm(45)) ,
                "imgd9": InlineImage(doc,foto3[9],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act10": fp_mes_actual[10] ,
                "fp_mes_ant10": fp_mes_anterior[10] ,
                "fp_mes10": fp_mes[10] ,
                "hp_mes_act10": hp_mes_actual[10] ,
                "hp_mes_ant10": hp_mes_anterior[10] ,
                "hp_mes10": hp_mes[10] ,
                "ea_mes_act10": ea_mes_actual[10] ,
                "ea_mes_ant10": ea_mes_anterior[10] ,
                "ea_mes10": ea_mes[10] ,
                "maxima_demanda10": maxima_demanda[10] ,
                "er_mes_act10": er_mes_actual[10] ,
                "er_mes_ant10": er_mes_anterior[10] ,
                "er_mes10": er_mes[10] ,
                "nom_tablero10":nom_tablero[10] ,
                "imgab10": InlineImage(doc,foto1[10],height=Mm(35), width=Mm(45)) ,
                "imgc10": InlineImage(doc,foto2[10],height=Mm(35), width=Mm(45)) ,
                "imgd10": InlineImage(doc,foto3[10],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act11": fp_mes_actual[11] ,
                "fp_mes_ant11": fp_mes_anterior[11] ,
                "fp_mes11": fp_mes[11] ,
                "hp_mes_act11": hp_mes_actual[11] ,
                "hp_mes_ant11": hp_mes_anterior[11] ,
                "hp_mes11": hp_mes[11] ,
                "ea_mes_act11": ea_mes_actual[11] ,
                "ea_mes_ant11": ea_mes_anterior[11] ,
                "ea_mes11": ea_mes[11] ,
                "maxima_demanda11": maxima_demanda[11] ,
                "er_mes_act11": er_mes_actual[11] ,
                "er_mes_ant11": er_mes_anterior[11] ,
                "er_mes11": er_mes[11] ,
                "nom_tablero11":nom_tablero[11] ,
                "imgab11": InlineImage(doc,foto1[11],height=Mm(35), width=Mm(45)) ,
                "imgc11": InlineImage(doc,foto2[11],height=Mm(35), width=Mm(45)) ,
                "imgd11": InlineImage(doc,foto3[11],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 12")
            if cont_tableros == 13:
                doc = DocxTemplate("templates/template13.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act5": fp_mes_actual[5] ,
                "fp_mes_ant5": fp_mes_anterior[5] ,
                "fp_mes5": fp_mes[5],
                "hp_mes_act5": hp_mes_actual[5] ,
                "hp_mes_ant5": hp_mes_anterior[5] ,
                "hp_mes5": hp_mes[5] ,
                "ea_mes_act5": ea_mes_actual[5] ,
                "ea_mes_ant5": ea_mes_anterior[5] ,
                "ea_mes5": ea_mes[5] ,
                "maxima_demanda5": maxima_demanda[5] ,
                "er_mes_act5": er_mes_actual[5] ,
                "er_mes_ant5": er_mes_anterior[5] ,
                "er_mes5": er_mes[5] ,
                "nom_tablero5":nom_tablero[5] ,
                "imgab5": InlineImage(doc,foto1[5],height=Mm(35), width=Mm(45)) ,
                "imgc5": InlineImage(doc,foto2[5],height=Mm(35), width=Mm(45)) ,
                "imgd5": InlineImage(doc,foto3[5],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act6": fp_mes_actual[6] ,
                "fp_mes_ant6": fp_mes_anterior[6] ,
                "fp_mes6": fp_mes[6] ,
                "hp_mes_act6": hp_mes_actual[6] ,
                "hp_mes_ant6": hp_mes_anterior[6] ,
                "hp_mes6": hp_mes[6] ,
                "ea_mes_act6": ea_mes_actual[6] ,
                "ea_mes_ant6": ea_mes_anterior[6] ,
                "ea_mes6": ea_mes[6] ,
                "maxima_demanda6": maxima_demanda[6] ,
                "er_mes_act6": er_mes_actual[6] ,
                "er_mes_ant6": er_mes_anterior[6] ,
                "er_mes6": er_mes[6] ,
                "nom_tablero6":nom_tablero[6] ,
                "imgab6": InlineImage(doc,foto1[6],height=Mm(35), width=Mm(45)) ,
                "imgc6": InlineImage(doc,foto2[6],height=Mm(35), width=Mm(45)) ,
                "imgd6": InlineImage(doc,foto3[6],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act7": fp_mes_actual[7] ,
                "fp_mes_ant7": fp_mes_anterior[7] ,
                "fp_mes7": fp_mes[7] ,
                "hp_mes_act7": hp_mes_actual[7] ,
                "hp_mes_ant7": hp_mes_anterior[7] ,
                "hp_mes7": hp_mes[7] ,
                "ea_mes_act7": ea_mes_actual[7] ,
                "ea_mes_ant7": ea_mes_anterior[7] ,
                "ea_mes7": ea_mes[7] ,
                "maxima_demanda7": maxima_demanda[7] ,
                "er_mes_act7": er_mes_actual[7] ,
                "er_mes_ant7": er_mes_anterior[7] ,
                "er_mes7": er_mes[7] ,
                "nom_tablero7":nom_tablero[7] ,
                "imgab7": InlineImage(doc,foto1[7],height=Mm(35), width=Mm(45)) ,
                "imgc7": InlineImage(doc,foto2[7],height=Mm(35), width=Mm(45)) ,
                "imgd7": InlineImage(doc,foto3[7],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act8": fp_mes_actual[8] ,
                "fp_mes_ant8": fp_mes_anterior[8] ,
                "fp_mes8": fp_mes[8] ,
                "hp_mes_act8": hp_mes_actual[8] ,
                "hp_mes_ant8": hp_mes_anterior[8] ,
                "hp_mes8": hp_mes[8] ,
                "ea_mes_act8": ea_mes_actual[8] ,
                "ea_mes_ant8": ea_mes_anterior[8] ,
                "ea_mes8": ea_mes[8] ,
                "maxima_demanda8": maxima_demanda[8] ,
                "er_mes_act8": er_mes_actual[8] ,
                "er_mes_ant8": er_mes_anterior[8] ,
                "er_mes8": er_mes[8] ,
                "nom_tablero8":nom_tablero[8] ,
                "imgab8": InlineImage(doc,foto1[8],height=Mm(35), width=Mm(45)) ,
                "imgc8": InlineImage(doc,foto2[8],height=Mm(35), width=Mm(45)) ,
                "imgd8": InlineImage(doc,foto3[8],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act9": fp_mes_actual[9] ,
                "fp_mes_ant9": fp_mes_anterior[9] ,
                "fp_mes9": fp_mes[9] ,
                "hp_mes_act9": hp_mes_actual[9] ,
                "hp_mes_ant9": hp_mes_anterior[9] ,
                "hp_mes9": hp_mes[9] ,
                "ea_mes_act9": ea_mes_actual[9] ,
                "ea_mes_ant9": ea_mes_anterior[9] ,
                "ea_mes9": ea_mes[9] ,
                "maxima_demanda9": maxima_demanda[9] ,
                "er_mes_act9": er_mes_actual[9] ,
                "er_mes_ant9": er_mes_anterior[9] ,
                "er_mes9": er_mes[9] ,
                "nom_tablero9":nom_tablero[9] ,
                "imgab9": InlineImage(doc,foto1[9],height=Mm(35), width=Mm(45)) ,
                "imgc9": InlineImage(doc,foto2[9],height=Mm(35), width=Mm(45)) ,
                "imgd9": InlineImage(doc,foto3[9],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act10": fp_mes_actual[10] ,
                "fp_mes_ant10": fp_mes_anterior[10] ,
                "fp_mes10": fp_mes[10] ,
                "hp_mes_act10": hp_mes_actual[10] ,
                "hp_mes_ant10": hp_mes_anterior[10] ,
                "hp_mes10": hp_mes[10] ,
                "ea_mes_act10": ea_mes_actual[10] ,
                "ea_mes_ant10": ea_mes_anterior[10] ,
                "ea_mes10": ea_mes[10] ,
                "maxima_demanda10": maxima_demanda[10] ,
                "er_mes_act10": er_mes_actual[10] ,
                "er_mes_ant10": er_mes_anterior[10] ,
                "er_mes10": er_mes[10] ,
                "nom_tablero10":nom_tablero[10] ,
                "imgab10": InlineImage(doc,foto1[10],height=Mm(35), width=Mm(45)) ,
                "imgc10": InlineImage(doc,foto2[10],height=Mm(35), width=Mm(45)) ,
                "imgd10": InlineImage(doc,foto3[10],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act11": fp_mes_actual[11] ,
                "fp_mes_ant11": fp_mes_anterior[11] ,
                "fp_mes11": fp_mes[11] ,
                "hp_mes_act11": hp_mes_actual[11] ,
                "hp_mes_ant11": hp_mes_anterior[11] ,
                "hp_mes11": hp_mes[11] ,
                "ea_mes_act11": ea_mes_actual[11] ,
                "ea_mes_ant11": ea_mes_anterior[11] ,
                "ea_mes11": ea_mes[11] ,
                "maxima_demanda11": maxima_demanda[11] ,
                "er_mes_act11": er_mes_actual[11] ,
                "er_mes_ant11": er_mes_anterior[11] ,
                "er_mes11": er_mes[11] ,
                "nom_tablero11":nom_tablero[11] ,
                "imgab11": InlineImage(doc,foto1[11],height=Mm(35), width=Mm(45)) ,
                "imgc11": InlineImage(doc,foto2[11],height=Mm(35), width=Mm(45)) ,
                "imgd11": InlineImage(doc,foto3[11],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act12": fp_mes_actual[12] ,
                "fp_mes_ant12": fp_mes_anterior[12] ,
                "fp_mes12": fp_mes[12] ,
                "hp_mes_act12": hp_mes_actual[12] ,
                "hp_mes_ant12": hp_mes_anterior[12] ,
                "hp_mes12": hp_mes[12] ,
                "ea_mes_act12": ea_mes_actual[12] ,
                "ea_mes_ant12": ea_mes_anterior[12] ,
                "ea_mes12": ea_mes[12] ,
                "maxima_demanda12": maxima_demanda[12] ,
                "er_mes_act12": er_mes_actual[12] ,
                "er_mes_ant12": er_mes_anterior[12] ,
                "er_mes12": er_mes[12] ,
                "nom_tablero12":nom_tablero[12] ,
                "imgab12": InlineImage(doc,foto1[12],height=Mm(35), width=Mm(45)) ,
                "imgc12": InlineImage(doc,foto2[12],height=Mm(35), width=Mm(45)) ,
                "imgd12": InlineImage(doc,foto3[12],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 13")
            if cont_tableros == 14:
                doc = DocxTemplate("templates/template14.docx")
                context = {
                "informe_tecnico" : informe_tecnico,
                "tablero" : tablero,
                "fecha_lectura": fecha_lectura,
                "fecha_emision": fecha_emision,
                "cliente": cliente,
                "area": area,
                "p1":  p1 ,
                "p2":  p2 ,
                "p3":  p3 ,
                "p10": round((cargo_potencia_activa_generacion_presente_punta),2) ,
                "p11": round((cargo_potencia_activa_generacion_presente_fuera_punta),2) ,
                "p12": round((cargo_potencia_activa_redes_presente_punta),2) ,
                "p13": round((cargo_potencia_activa_redes_presente_fuera_punta),2) ,
                "cargo_fijo_mensual": cargo_fijo_mensual,
                "cargo_energia_activa_punta": cargo_energia_activa_punta ,
                "cargo_energia_activa_fuera_punta": cargo_energia_activa_fuera_punta ,
                "cargo_potencia_activa_generacion_presente_punta": cargo_potencia_activa_generacion_presente_punta ,
                "cargo_potencia_activa_generacion_presente_fuera_punta": cargo_potencia_activa_generacion_presente_fuera_punta ,
                "cargo_potencia_activa_redes_presente_punta": cargo_potencia_activa_redes_presente_punta ,
                "cargo_potencia_activa_redes_presente_fuera_punta": cargo_potencia_activa_redes_presente_fuera_punta ,
                "cargo_energia_reactiva_exc_30": cargo_energia_reactiva_exc_30 ,

                "fp_mes_act0": fp_mes_actual[0] ,
                "fp_mes_ant0": fp_mes_anterior[0] ,
                "fp_mes0": fp_mes[0] ,
                "hp_mes_act0": hp_mes_actual[0] ,
                "hp_mes_ant0": hp_mes_anterior[0] ,
                "hp_mes0": hp_mes[0] ,
                "ea_mes_act0": ea_mes_actual[0] ,
                "ea_mes_ant0": ea_mes_anterior[0] ,
                "ea_mes0": ea_mes[0] ,
                "maxima_demanda0": maxima_demanda[0] ,
                "er_mes_act0": er_mes_actual[0] ,
                "er_mes_ant0": er_mes_anterior[0] ,
                "er_mes0": er_mes[0] ,
                "nom_tablero0":nom_tablero[0] ,
                "imgab0": InlineImage(doc,foto1[0],height=Mm(35), width=Mm(45)) ,
                "imgc0": InlineImage(doc,foto2[0],height=Mm(35), width=Mm(45)) ,
                "imgd0": InlineImage(doc,foto3[0],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act1": fp_mes_actual[1] ,
                "fp_mes_ant1": fp_mes_anterior[1] ,
                "fp_mes1": fp_mes[1] ,
                "hp_mes_act1": hp_mes_actual[1] ,
                "hp_mes_ant1": hp_mes_anterior[1] ,
                "hp_mes1": hp_mes[1] ,
                "ea_mes_act1": ea_mes_actual[1] ,
                "ea_mes_ant1": ea_mes_anterior[1] ,
                "ea_mes1": ea_mes[1] ,
                "maxima_demanda1": maxima_demanda[1] ,
                "er_mes_act1": er_mes_actual[1] ,
                "er_mes_ant1": er_mes_anterior[1] ,
                "er_mes1": er_mes[1] ,
                "nom_tablero1":nom_tablero[1] ,
                "imgab1": InlineImage(doc,foto1[1],height=Mm(35), width=Mm(45)) ,
                "imgc1": InlineImage(doc,foto2[1],height=Mm(35), width=Mm(45)) ,
                "imgd1": InlineImage(doc,foto3[1],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act2": fp_mes_actual[2] ,
                "fp_mes_ant2": fp_mes_anterior[2] ,
                "fp_mes2": fp_mes[2] ,
                "hp_mes_act2": hp_mes_actual[2] ,
                "hp_mes_ant2": hp_mes_anterior[2] ,
                "hp_mes2": hp_mes[2] ,
                "ea_mes_act2": ea_mes_actual[2] ,
                "ea_mes_ant2": ea_mes_anterior[2] ,
                "ea_mes2": ea_mes[2] ,
                "maxima_demanda2": maxima_demanda[2] ,
                "er_mes_act2": er_mes_actual[2] ,
                "er_mes_ant2": er_mes_anterior[2] ,
                "er_mes2": er_mes[2] ,
                "nom_tablero2":nom_tablero[2] ,
                "imgab2": InlineImage(doc,foto1[2],height=Mm(35), width=Mm(45)) ,
                "imgc2": InlineImage(doc,foto2[2],height=Mm(35), width=Mm(45)) ,
                "imgd2": InlineImage(doc,foto3[2],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act3": fp_mes_actual[3] ,
                "fp_mes_ant3": fp_mes_anterior[3] ,
                "fp_mes3": fp_mes[3] ,
                "hp_mes_act3": hp_mes_actual[3] ,
                "hp_mes_ant3": hp_mes_anterior[3] ,
                "hp_mes3": hp_mes[3] ,
                "ea_mes_act3": ea_mes_actual[3] ,
                "ea_mes_ant3": ea_mes_anterior[3] ,
                "ea_mes3": ea_mes[3] ,
                "maxima_demanda3": maxima_demanda[3] ,
                "er_mes_act3": er_mes_actual[3] ,
                "er_mes_ant3": er_mes_anterior[3] ,
                "er_mes3": er_mes[3] ,
                "nom_tablero3":nom_tablero[3] ,
                "imgab3": InlineImage(doc,foto1[3],height=Mm(35), width=Mm(45)) ,
                "imgc3": InlineImage(doc,foto2[3],height=Mm(35), width=Mm(45)) ,
                "imgd3": InlineImage(doc,foto3[3],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act4": fp_mes_actual[4] ,
                "fp_mes_ant4": fp_mes_anterior[4] ,
                "fp_mes4": fp_mes[4] ,
                "hp_mes_act4": hp_mes_actual[4] ,
                "hp_mes_ant4": hp_mes_anterior[4] ,
                "hp_mes4": hp_mes[4] ,
                "ea_mes_act4": ea_mes_actual[4] ,
                "ea_mes_ant4": ea_mes_anterior[4] ,
                "ea_mes4": ea_mes[4] ,
                "maxima_demanda4": maxima_demanda[4] ,
                "er_mes_act4": er_mes_actual[4] ,
                "er_mes_ant4": er_mes_anterior[4] ,
                "er_mes4": er_mes[4] ,
                "nom_tablero4":nom_tablero[4] ,
                "imgab4": InlineImage(doc,foto1[4],height=Mm(35), width=Mm(45)) ,
                "imgc4": InlineImage(doc,foto2[4],height=Mm(35), width=Mm(45)) ,
                "imgd4": InlineImage(doc,foto3[4],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act5": fp_mes_actual[5] ,
                "fp_mes_ant5": fp_mes_anterior[5] ,
                "fp_mes5": fp_mes[5],
                "hp_mes_act5": hp_mes_actual[5] ,
                "hp_mes_ant5": hp_mes_anterior[5] ,
                "hp_mes5": hp_mes[5] ,
                "ea_mes_act5": ea_mes_actual[5] ,
                "ea_mes_ant5": ea_mes_anterior[5] ,
                "ea_mes5": ea_mes[5] ,
                "maxima_demanda5": maxima_demanda[5] ,
                "er_mes_act5": er_mes_actual[5] ,
                "er_mes_ant5": er_mes_anterior[5] ,
                "er_mes5": er_mes[5] ,
                "nom_tablero5":nom_tablero[5] ,
                "imgab5": InlineImage(doc,foto1[5],height=Mm(35), width=Mm(45)) ,
                "imgc5": InlineImage(doc,foto2[5],height=Mm(35), width=Mm(45)) ,
                "imgd5": InlineImage(doc,foto3[5],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act6": fp_mes_actual[6] ,
                "fp_mes_ant6": fp_mes_anterior[6] ,
                "fp_mes6": fp_mes[6] ,
                "hp_mes_act6": hp_mes_actual[6] ,
                "hp_mes_ant6": hp_mes_anterior[6] ,
                "hp_mes6": hp_mes[6] ,
                "ea_mes_act6": ea_mes_actual[6] ,
                "ea_mes_ant6": ea_mes_anterior[6] ,
                "ea_mes6": ea_mes[6] ,
                "maxima_demanda6": maxima_demanda[6] ,
                "er_mes_act6": er_mes_actual[6] ,
                "er_mes_ant6": er_mes_anterior[6] ,
                "er_mes6": er_mes[6] ,
                "nom_tablero6":nom_tablero[6] ,
                "imgab6": InlineImage(doc,foto1[6],height=Mm(35), width=Mm(45)) ,
                "imgc6": InlineImage(doc,foto2[6],height=Mm(35), width=Mm(45)) ,
                "imgd6": InlineImage(doc,foto3[6],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act7": fp_mes_actual[7] ,
                "fp_mes_ant7": fp_mes_anterior[7] ,
                "fp_mes7": fp_mes[7] ,
                "hp_mes_act7": hp_mes_actual[7] ,
                "hp_mes_ant7": hp_mes_anterior[7] ,
                "hp_mes7": hp_mes[7] ,
                "ea_mes_act7": ea_mes_actual[7] ,
                "ea_mes_ant7": ea_mes_anterior[7] ,
                "ea_mes7": ea_mes[7] ,
                "maxima_demanda7": maxima_demanda[7] ,
                "er_mes_act7": er_mes_actual[7] ,
                "er_mes_ant7": er_mes_anterior[7] ,
                "er_mes7": er_mes[7] ,
                "nom_tablero7":nom_tablero[7] ,
                "imgab7": InlineImage(doc,foto1[7],height=Mm(35), width=Mm(45)) ,
                "imgc7": InlineImage(doc,foto2[7],height=Mm(35), width=Mm(45)) ,
                "imgd7": InlineImage(doc,foto3[7],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act8": fp_mes_actual[8] ,
                "fp_mes_ant8": fp_mes_anterior[8] ,
                "fp_mes8": fp_mes[8] ,
                "hp_mes_act8": hp_mes_actual[8] ,
                "hp_mes_ant8": hp_mes_anterior[8] ,
                "hp_mes8": hp_mes[8] ,
                "ea_mes_act8": ea_mes_actual[8] ,
                "ea_mes_ant8": ea_mes_anterior[8] ,
                "ea_mes8": ea_mes[8] ,
                "maxima_demanda8": maxima_demanda[8] ,
                "er_mes_act8": er_mes_actual[8] ,
                "er_mes_ant8": er_mes_anterior[8] ,
                "er_mes8": er_mes[8] ,
                "nom_tablero8":nom_tablero[8] ,
                "imgab8": InlineImage(doc,foto1[8],height=Mm(35), width=Mm(45)) ,
                "imgc8": InlineImage(doc,foto2[8],height=Mm(35), width=Mm(45)) ,
                "imgd8": InlineImage(doc,foto3[8],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act9": fp_mes_actual[9] ,
                "fp_mes_ant9": fp_mes_anterior[9] ,
                "fp_mes9": fp_mes[9] ,
                "hp_mes_act9": hp_mes_actual[9] ,
                "hp_mes_ant9": hp_mes_anterior[9] ,
                "hp_mes9": hp_mes[9] ,
                "ea_mes_act9": ea_mes_actual[9] ,
                "ea_mes_ant9": ea_mes_anterior[9] ,
                "ea_mes9": ea_mes[9] ,
                "maxima_demanda9": maxima_demanda[9] ,
                "er_mes_act9": er_mes_actual[9] ,
                "er_mes_ant9": er_mes_anterior[9] ,
                "er_mes9": er_mes[9] ,
                "nom_tablero9":nom_tablero[9] ,
                "imgab9": InlineImage(doc,foto1[9],height=Mm(35), width=Mm(45)) ,
                "imgc9": InlineImage(doc,foto2[9],height=Mm(35), width=Mm(45)) ,
                "imgd9": InlineImage(doc,foto3[9],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act10": fp_mes_actual[10] ,
                "fp_mes_ant10": fp_mes_anterior[10] ,
                "fp_mes10": fp_mes[10] ,
                "hp_mes_act10": hp_mes_actual[10] ,
                "hp_mes_ant10": hp_mes_anterior[10] ,
                "hp_mes10": hp_mes[10] ,
                "ea_mes_act10": ea_mes_actual[10] ,
                "ea_mes_ant10": ea_mes_anterior[10] ,
                "ea_mes10": ea_mes[10] ,
                "maxima_demanda10": maxima_demanda[10] ,
                "er_mes_act10": er_mes_actual[10] ,
                "er_mes_ant10": er_mes_anterior[10] ,
                "er_mes10": er_mes[10] ,
                "nom_tablero10":nom_tablero[10] ,
                "imgab10": InlineImage(doc,foto1[10],height=Mm(35), width=Mm(45)) ,
                "imgc10": InlineImage(doc,foto2[10],height=Mm(35), width=Mm(45)) ,
                "imgd10": InlineImage(doc,foto3[10],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act11": fp_mes_actual[11] ,
                "fp_mes_ant11": fp_mes_anterior[11] ,
                "fp_mes11": fp_mes[11] ,
                "hp_mes_act11": hp_mes_actual[11] ,
                "hp_mes_ant11": hp_mes_anterior[11] ,
                "hp_mes11": hp_mes[11] ,
                "ea_mes_act11": ea_mes_actual[11] ,
                "ea_mes_ant11": ea_mes_anterior[11] ,
                "ea_mes11": ea_mes[11] ,
                "maxima_demanda11": maxima_demanda[11] ,
                "er_mes_act11": er_mes_actual[11] ,
                "er_mes_ant11": er_mes_anterior[11] ,
                "er_mes11": er_mes[11] ,
                "nom_tablero11":nom_tablero[11] ,
                "imgab11": InlineImage(doc,foto1[11],height=Mm(35), width=Mm(45)) ,
                "imgc11": InlineImage(doc,foto2[11],height=Mm(35), width=Mm(45)) ,
                "imgd11": InlineImage(doc,foto3[11],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act12": fp_mes_actual[12] ,
                "fp_mes_ant12": fp_mes_anterior[12] ,
                "fp_mes12": fp_mes[12] ,
                "hp_mes_act12": hp_mes_actual[12] ,
                "hp_mes_ant12": hp_mes_anterior[12] ,
                "hp_mes12": hp_mes[12] ,
                "ea_mes_act12": ea_mes_actual[12] ,
                "ea_mes_ant12": ea_mes_anterior[12] ,
                "ea_mes12": ea_mes[12] ,
                "maxima_demanda12": maxima_demanda[12] ,
                "er_mes_act12": er_mes_actual[12] ,
                "er_mes_ant12": er_mes_anterior[12] ,
                "er_mes12": er_mes[12] ,
                "nom_tablero12":nom_tablero[12] ,
                "imgab12": InlineImage(doc,foto1[12],height=Mm(35), width=Mm(45)) ,
                "imgc12": InlineImage(doc,foto2[12],height=Mm(35), width=Mm(45)) ,
                "imgd12": InlineImage(doc,foto3[12],height=Mm(35), width=Mm(45)) ,
                "fp_mes_act13": fp_mes_actual[13] ,
                "fp_mes_ant13": fp_mes_anterior[13] ,
                "fp_mes13": fp_mes[13] ,
                "hp_mes_act13": hp_mes_actual[13] ,
                "hp_mes_ant13": hp_mes_anterior[13] ,
                "hp_mes13": hp_mes[13] ,
                "ea_mes_act13": ea_mes_actual[13] ,
                "ea_mes_ant13": ea_mes_anterior[13] ,
                "ea_mes13": ea_mes[13] ,
                "maxima_demanda13": maxima_demanda[13] ,
                "er_mes_act13": er_mes_actual[13] ,
                "er_mes_ant13": er_mes_anterior[13] ,
                "er_mes13": er_mes[13] ,
                "nom_tablero13":nom_tablero[13] ,
                "imgab13": InlineImage(doc,foto1[13],height=Mm(35), width=Mm(45)) ,
                "imgc13": InlineImage(doc,foto2[13],height=Mm(35), width=Mm(45)) ,
                "imgd13": InlineImage(doc,foto3[13],height=Mm(35), width=Mm(45)) ,
                
                "hp_mestotal": hp_mestotal ,
                "maxima_demandatotal": maxima_demandatotal ,
                "fp_mestotal": fp_mestotal  ,
                "ea_mestotal": ea_mestotal ,
                "er_mestotal": er_mestotal ,
                "er_mestotalfac": er_mestotalfac,
                "dias": dias,
                "horas_punta": horas_punta ,
                "promedio":promedio ,
                "texto": texto,
                "texto2": texto2,
                "cali_tari": cali_tari,
                "cargo_eapp": cargo_eapp,
                "cargo_eafp": cargo_eafp,
                "operacion1": operacion1 ,
                "operacion2": operacion2 ,
                "operacion3": operacion3 ,
                "operacion4": operacion4 ,
                "v2": v2,
                "v3": v3,
                "v4": v4,
                "v5": v5,
                "v6": v6,
                "v7": v7,
                "cargo_er_exc30": cargo_er_exc30,
                "cargo_pagpp":  cargo_pagpp,
                "cargo_pagfp":  cargo_pagfp,
                "cargo_parpp":  cargo_parpp,
                "cargo_parfp":  cargo_parfp,
                "hp_mestotal": hp_mestotal,
                "fp_mestotal" : fp_mestotal,
                "subtotal": round(subtotal,2),
                "conigv": round(conigv,2),
                "total_final": round(total_final,2),

                "nombre_mes1": nom_mes1,
                "cantidad_mes1": round(meses[0],2),
                "nombre_mes2": nom_mes2,
                "cantidad_mes2": round(meses[1],2),
                "nombre_mes3": nom_mes3,
                "cantidad_mes3": round(meses[2],2),
                "nombre_mes4": nom_mes4,
                "cantidad_mes4": round(meses[3],2),
                "nombre_mes5": nom_mes5,
                "cantidad_mes5": round(meses[4],2),
                "nombre_mes6": nom_mes6,
                "cantidad_mes6": round(meses[5],2),
                "anio": anio
                }
                print("template 14")
            
            doc.render(context)
            nombrefinal="INFORME TECNICO N° "+informe_tecnico+".docx"
            print(nombrefinal)
            doc.save(nombrefinal)
            mensaje3(nombrefinal)
            reiniciar()


        btnfinal1=Button(framefinal,text="Guardar",command= asignar,  font=("Arial", 12),cursor="hand2", relief="raised", borderwidth=4)
        btnfinal1.grid(row=0,column=1,padx=30,pady=2)
        btnfinal1.config(bg="#ecefee")

        btnfinal2=Button(framefinal,text="Exportar", command= Exportar,font=("Arial", 12),cursor="hand2", relief="raised", borderwidth=4)
        btnfinal2.grid(row=0,column=2,padx=30,pady=2)
        btnfinal2.config(bg="#ecefee")
        btnfinal2.configure(state='disabled')

        btnfinal3=Button(framefinal,text="Limpiar", command= reiniciar,font=("Arial", 12),cursor="hand2", relief="raised", borderwidth=4)
        btnfinal3.grid(row=0,column=0,padx=30, pady=2)
        btnfinal3.config(bg="#ecefee")

        mycanvas.create_window((0,0), window=miframe, anchor="nw")

        wraper1.pack(fill="both", expand="yes", padx=0, pady=0)

        raiz.geometry("945x810")
        raiz.resizable(False,True)

        raiz.mainloop()