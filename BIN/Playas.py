import os
import pandas as pd
import openpyxl
import numpy as np
from tkinter.filedialog import askdirectory
from tkinter import messagebox
import xlsxwriter

def ConsolidarExcels():

    # Seleccionar la carpeta donde se encuentran los archivos
    path = askdirectory(title='Seleccionar carpeta donde se encuentran los archivos')

    # Listar todos los archivos de la carpeta 'Archivos'
    archivos = os.listdir(path)

    # Filtrar los que no sean .xlsx
    archivos = [archivo for archivo in archivos if archivo[-4:] == 'xlsx']

    # Filtrar los que sean 'Consolidado.xlsx'
    archivos = [archivo for archivo in archivos if archivo != 'Consolidado.xlsx']

    # Crear un DataFrame vacío
    ConsolidadoTPT = pd.DataFrame()
    ConsolidadoAPT = pd.DataFrame()

    # Iterar sobre los archivos
    for archivo in archivos:
        # Leer la celda 'C2' de la Primer hoja y almacenarlo en la Variable 'Descripción de playa'
        descripcion = openpyxl.load_workbook(f'{path}/{archivo}').worksheets[0]['C2'].value
        # Sepeara la descripción por '-' y quedarse con la segunda parte
        descripcion = descripcion.split('-')[1]
        # Eliminar los espacios en blanco al inicio y al final
        descripcion = descripcion.strip()

        # Leer los datos de la hoja 'ListadoTanques' y almacenarlo en la Variable 'Tanques' 
        tanques = pd.read_excel(f'{path}/{archivo}', sheet_name='ListadoTanques')
        # Agregar una columna con la descripción de la playa al DataFrame 'Tanques'
        tanques['Descripción de playa'] = descripcion

        # # Leer los datos de la hoja 'EstadoAforadores' y almacenarlo en la Variable 'Aforadores'
        # aforadores = pd.read_excel(f'Archivos/{archivo}', sheet_name='EstadoAforadores')
        # # Agregar una columna con la descripción de la playa al DataFrame 'Aforadores'
        # aforadores['Descripción de playa'] = descripcion

        # Crear una tabla pivote de tanques por 'DescripcionArticulo' donde se sumen las columnas 'Descarga' , 'FacturaRemito' , 'Medicion' , 'Vendido' , 'StockActual' , 'Vacio' , 'StockAnterior' , 'Diferencia' , 'PorcentajeDiferencia' , 'VentaPromedio'
        tanquesPT = tanques.pivot_table(index=['Descripción de playa' , 'DescripcionArticulo'], values=['Descarga', 'FacturaRemito', 'Medicion', 'Vendido', 'StockActual', 'Vacio', 'StockAnterior', 'Diferencia'], aggfunc=np.sum)

        # # Crear una tabla pivote de aforadores por 'DescripcionArticulo' donde se sumen las columna de 'Despachado'
        # aforadoresPT = aforadores.pivot_table(index=['Descripción de playa' , 'DescripcionArticulo'], values=['Despachado'], aggfunc=np.sum)

        # Consolidar los datos de la tabla pivote en el DataFrame 'ConsolidadoTPT'
        ConsolidadoTPT = pd.concat([ConsolidadoTPT, tanquesPT])

        # # Consolidar los datos de la tabla pivote en el DataFrame 'ConsolidadoAPT'
        # ConsolidadoAPT = pd.concat([ConsolidadoAPT, aforadoresPT])

    # Seleccionar la columna 'StockActual'
    Actual = ConsolidadoTPT['StockActual']
    # Transformar indice en columnas
    Actual = Actual.reset_index()
    # Transformar la tabla para que las 'DescripcionArticulo' sean columnas, los 'StockActual' sean los valores y los 'Descripción de playa' sean las filas
    Actual = Actual.pivot_table(index=['Descripción de playa'], columns=['DescripcionArticulo'], values=['StockActual'] , aggfunc=np.sum)
    # Agregar una fila con la suma de los valores de cada columna
    Actual.loc['Total'] = Actual.sum(axis=0)


    # Seleccionar la columna 'StockAnterior'
    Anterior = ConsolidadoTPT['StockAnterior']
    # Transformar indice en columnas
    Anterior = Anterior.reset_index()
    # Transformar la tabla para que las 'DescripcionArticulo' sean columnas, los 'StockAnterior' sean los valores y los 'Descripción de playa' sean las filas
    Anterior = Anterior.pivot_table(index=['Descripción de playa'], columns=['DescripcionArticulo'], values=['StockAnterior'] , aggfunc=np.sum)
    # Agregar una fila con la suma de los valores de cada columna
    Anterior.loc['Total'] = Anterior.sum(axis=0)


    # Seleccionar la columna 'Vendido'
    Vendido = ConsolidadoTPT['Vendido']
    # Transformar indice en columnas
    Vendido = Vendido.reset_index()
    # Transformar la tabla para que las 'DescripcionArticulo' sean columnas, los 'Vendido' sean los valores y los 'Descripción de playa' sean las filas
    Vendido = Vendido.pivot_table(index=['Descripción de playa'], columns=['DescripcionArticulo'], values=['Vendido'] , aggfunc=np.sum)
    # Agregar una fila con la suma de los valores de cada columna
    Vendido.loc['Total'] = Vendido.sum(axis=0)


    # aforadoresPT_Exportar = ConsolidadoAPT['Despachado']
    # aforadoresPT_Exportar = aforadoresPT_Exportar.reset_index()
    # aforadoresPT_Exportar = ConsolidadoAPT.pivot_table(index=['Descripción de playa'], columns=['DescripcionArticulo'], values=['Despachado'] , aggfunc=np.sum)

    # Exportar Actual, Anterior y Vendido a un archivo Excel
    # with pd.ExcelWriter('ConsolidadoPT.xlsx') as writer:
    #     Actual.to_excel(writer, sheet_name='Actual')
    #     Anterior.to_excel(writer, sheet_name='Anterior')
    #     Vendido.to_excel(writer, sheet_name='Vendido')
        # aforadoresPT_Exportar.to_excel(writer, sheet_name='Aforadores')

    # Exportar los 4 dataframes a un a misma hoja de un archivo Excel donde primero se exporte Anterior, luego Actual, luego Vendido y por último Aforadores. Separados por una fila en blanco
    # with pd.ExcelWriter(( path + "/" +'Consolidado.xlsx')) as writer:
    #     Anterior.to_excel(writer, sheet_name='Consolidado', startrow=0)
    #     Actual.to_excel(writer, sheet_name='Consolidado', startrow=Anterior.shape[0]+4)
    #     Vendido.to_excel(writer, sheet_name='Consolidado', startrow=Anterior.shape[0]+Actual.shape[0]+8)
        # aforadoresPT_Exportar.to_excel(writer, sheet_name='Consolidado', startrow=Anterior.shape[0]+Actual.shape[0]+Vendido.shape[0]+12)

    # Exportar los 4 dataframes a un a misma hoja de un archivo Excel donde primero se exporte Anterior, luego Actual, luego Vendido y por último Aforadores. Separados por una fila en blanco con formato de tabla con estilo 'Table Style Medium 2'
    with pd.ExcelWriter((path + "/" + 'Consolidado.xlsx')) as writer:
        Anterior.to_excel(writer, sheet_name='Consolidado', startrow=0)
        Actual.to_excel(writer, sheet_name='Consolidado', startrow=Anterior.shape[0]+4)
        Vendido.to_excel(writer, sheet_name='Consolidado', startrow=Anterior.shape[0]+Actual.shape[0]+8)

        # Crear un objeto de tipo workbook
        workbook = writer.book
        # Crear un objeto de tipo worksheet
        worksheet = writer.sheets['Consolidado']
        # Crear un objeto de tipo formato
        format = workbook.add_format({'border': 1 , 'num_format':'#,##0.00'} )
        format2 = workbook.add_format({'border': 1 , 'bg_color': '#D3D3D3' , 'num_format':'#,##0.00'} )
        format4 = workbook.add_format({'border': 1 , 'bg_color': '#203151' , 'num_format':'#,##0.00' , 'bold':True , 'font_color': '#FFFFFF'} )
        #Crear un formato donde el color de fondo sea Azul oscuro, las letras sean blancas y el texto este centrado y negrita para los titulos de las columnas
        format3 = workbook.add_format({'bg_color': '#0F243E', 'font_color': '#FFFFFF', 'align': 'center', 'bold': True})


        # Aplicar format3 a los titulos de las columnas de la tabla de 'Anterior'
        worksheet.conditional_format(1, 0, 1, Anterior.shape[1], {'type': 'no_blanks', 'format': format3})
        # Aplicar format a las filas con la condición que el numero de fila sea impar de la tabla de 'Anterior'
        worksheet.conditional_format(2, 0, Anterior.shape[0]+2, Anterior.shape[1], {'type': 'formula', 'criteria': '=MOD(ROW(),2)=1', 'format': format})
        # Aplicar format2 a las filas con la condición que el numero de fila sea par de la tabla de 'Anterior'
        worksheet.conditional_format(2, 0, Anterior.shape[0]+2, Anterior.shape[1], {'type': 'formula', 'criteria': '=MOD(ROW(),2)=0', 'format': format2})
        # Aplicar format4 a la fila que contiene la suma de los valores de cada columna de la tabla de 'Anterior'
        worksheet.conditional_format(Anterior.shape[0]+2, 0, Anterior.shape[0]+2, Anterior.shape[1], {'type': 'no_blanks', 'format': format4})
        
        worksheet.conditional_format(Anterior.shape[0]+5, 0, Anterior.shape[0]+5, Anterior.shape[1], {'type': 'no_blanks', 'format': format3})
        worksheet.conditional_format(Anterior.shape[0]+6, 0, Anterior.shape[0]+Actual.shape[0]+5, Anterior.shape[1], {'type': 'formula', 'criteria': '=MOD(ROW(),2)=1', 'format': format})
        worksheet.conditional_format(Anterior.shape[0]+6, 0, Anterior.shape[0]+Actual.shape[0]+5, Anterior.shape[1], {'type': 'formula', 'criteria': '=MOD(ROW(),2)=0', 'format': format2})
        worksheet.conditional_format(Anterior.shape[0]+Actual.shape[0]+6, 0, Anterior.shape[0]+Actual.shape[0]+6, Anterior.shape[1], {'type': 'no_blanks', 'format': format4})

        worksheet.conditional_format(Anterior.shape[0]+Actual.shape[0]+9, 0, Anterior.shape[0]+Actual.shape[0]+9, Anterior.shape[1], {'type': 'no_blanks', 'format': format3})
        worksheet.conditional_format(Anterior.shape[0]+Actual.shape[0]+10, 0, Anterior.shape[0]+Actual.shape[0]+Vendido.shape[0]+10, Anterior.shape[1], {'type': 'formula', 'criteria': '=MOD(ROW(),2)=1', 'format': format})
        worksheet.conditional_format(Anterior.shape[0]+Actual.shape[0]+10, 0, Anterior.shape[0]+Actual.shape[0]+Vendido.shape[0]+10, Anterior.shape[1], {'type': 'formula', 'criteria': '=MOD(ROW(),2)=0', 'format': format2})
        worksheet.conditional_format(Anterior.shape[0]+Actual.shape[0]+Vendido.shape[0]+10, 0, Anterior.shape[0]+Actual.shape[0]+Vendido.shape[0]+10, Anterior.shape[1], {'type': 'no_blanks', 'format': format4})

        # Autoajustar el ancho de las columnas del archivo Excel con openpyxl
        # for column in range(0, Anterior.shape[1]):
        #     max_length = 0
        #     for row in range(0, Anterior.shape[0]+Actual.shape[0]+Vendido.shape[0]+14):
        #         if len(str(Anterior.iloc[row,column])) > max_length:
        #             max_length = len(str(Anterior.iloc[row,column]))
        #     worksheet.set_column(column, column, max_length+2)


    # Mostar ventana emergente con el mensaje 'Proceso finalizado'
    messagebox.showinfo(message='Proceso finalizado', title='Información') 


if __name__ == '__main__':
    ConsolidarExcels()
        