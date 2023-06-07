#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
import BIN.Playas as Playas 


class ModeloPygubuApp:
    def __init__(self, master=None):
        # build ui
        Toplevel_1 = tk.Tk() if master is None else tk.Toplevel(master)
        Toplevel_1.configure(
            background="#2e2e2e",
            cursor="arrow",
            height=250,
            width=290)
        Toplevel_1.iconbitmap("BIN/ABP-blanco-en-fondo-negro.ico")
        Toplevel_1.minsize(290, 325)
        Toplevel_1.overrideredirect("False")
        Toplevel_1.title("Cálculo de Stock")
        Label_3 = ttk.Label(Toplevel_1)
        self.img_ABPblancoenfondonegro111 = tk.PhotoImage(
            file="BIN/ABP blanco en sin fondo .png")
        Label_3.configure(
            background="#2e2e2e",
            image=self.img_ABPblancoenfondonegro111)
        Label_3.pack(side="top")
        Label_1 = ttk.Label(Toplevel_1)
        Label_1.configure(
            background="#2e2e2e",
            font="TkDefaultFont",
            foreground="#ffffff",
            justify="center",
            state="disabled",
            takefocus=False,
            text='Cáclulo de Stock de Combustibles en base a los Excels emitidos por el sistema CALDEN OIL.\n',
            wraplength=325)
        Label_1.pack(expand="true", side="top")
        Label_2 = ttk.Label(Toplevel_1)
        Label_2.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            text='por Agustín Bustos Piasentini\nhttps://www.Agustin-Bustos-Piasentini.com.ar/\n')
        Label_2.pack(expand="true", side="top")
        self.Selección_Archivos = ttk.Button(Toplevel_1)
        self.Selección_Archivos.configure(
            text='Selección de Carpeta con Archivos de Excel',
            command=Playas.ConsolidarExcels)
        self.Selección_Archivos.pack(expand="true", pady=4, side="top")

        # Main widget
        self.mainwindow = Toplevel_1

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = ModeloPygubuApp()
    app.run()
