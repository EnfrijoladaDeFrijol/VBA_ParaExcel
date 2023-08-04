Attribute VB_Name = "Módulo1"
Option Explicit

' Me falta documentar, queda pendiente

Sub Enviar_a_PowerPoint()
    ' Variables
    Dim ppApp As PowerPoint.Application
    Dim ppPresentacion As PowerPoint.Presentation
    Dim rutaPlantilla As String
    Dim rutaActual As String
    
    ' Obtenemos rutas
    rutaActual = ThisWorkbook.Path ' Dir actual
    rutaPlantilla = rutaActual & "\PresentacionTablagraf.pptx"
    
    ' Inicializamos la variable ppApp solo si es Nothing
    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If ppApp Is Nothing Then
        Set ppApp = New PowerPoint.Application
    End If
    
    ' Abrimos la presentación en PowerPoint
    On Error Resume Next
    Set ppPresentacion = ppApp.Presentations.Open(rutaPlantilla)
    On Error GoTo 0
    
    ' Validamos si existe tal presentación
    If ppPresentacion Is Nothing Then
        MsgBox "La plantilla " & rutaPlantilla & " no se encuentra en la ruta" ' Caso no existe
    Else
        ' Vemos que tenga diapositivas
        If ppPresentacion.Slides.Count = 0 Then ' No tiene diapositivas
            MsgBox "La plantilla está vacía " & rutaPlantilla & " está vacía"
        Else ' Si tiene diapositivas
            ' Para la gráfica
            ThisWorkbook.Sheets("Hoja1").ChartObjects("Gráfico 1").Chart.ChartArea.Copy ' Activamos y copiamos el gráfico
            ppPresentacion.Slides(1).Shapes.PasteSpecial ppPasteJPG ' Pegamos como jpg en la diapositiva 1
            
            ' Para la tabla
            ThisWorkbook.Sheets("Hoja2").Range("A1:B8").Copy ' Copiamos el rango en la hoja 2
            ppPresentacion.Slides(2).Shapes.PasteSpecial ppPasteEnhancedMetafile ' Pegamos como metaarchivo mejorado en la diapositiva 2
        End If
    End If

End Sub
