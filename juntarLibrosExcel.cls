VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ImportarDatos()
    Dim wsHojaDestino As Worksheet ' Hoja donde pondremos los demas libros
    Dim ruta, misArchivos As String
    Dim wbLibroOrigen As Workbook ' Variable para guardar libros
    Dim ultimaFilaOrigen As Long ' Para hacer dinámico la última fila de cada archivo
    Dim ultimaFilaDestino As Long ' Para que se copien y sobreescriban
    
    Set wsHojaDestino = ThisWorkbook.Worksheets("Juntado") ' Decimos que de este libro tome una hoja
    ruta = ThisWorkbook.Path & "/misArchivos/" ' Guardamos la ruta donde están los archivos
    
    ' Nombre de cada libro
    ' Con "Dir" obtiene la cadena de nombres de una ruta especifica
    misArchivos = Dir(ruta & "*.xlsx")
    
    ' MsgBox misArchivos ' Listamos archivos
    ' MsgBox Len(misArchivos) ' Para ver la cantidad de caracteres que tiene una cadena
    
    Do While Len(misArchivos) > 0
        Set wbLibroOrigen = Workbooks.Open(ruta & misArchivos) ' Corregimos "archivos" a "misArchivos"
        ' Obtiene las filas de manera dinámica
        ultimaFilaOrigen = wbLibroOrigen.Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
        ' Comienza a pegar desde la última fila para abajo
        ultimaFilaDestino = wsHojaDestino.Range("A" & Rows.Count).End(xlUp).Row + 1
        
        ' Proceso de copiado
        wbLibroOrigen.Worksheets(1).Range("A2:F" & ultimaFilaOrigen).Copy Destination:=wsHojaDestino.Range("A" & ultimaFilaDestino)
        ' Para cerrar los libros que se van chequeando
        wbLibroOrigen.Close SaveChanges:=False
        ' MsgBox misArchivos
        misArchivos = Dir()
    Loop

    MsgBox "Libros importados correctamente"

End Sub




