# Herramienta de registro de Eventos  

🎯 ¿Qué problema resuelve?

La necesidad de centralizar y estandarizar el registro de eventos para reducir tiempos de consolidación, minimizar errores en la captura de información 
y mejorar la disponibilidad de datos para el seguimiento.

🛠️ Herramientas: Excel | Visual Basic

📈 Resultados: Herramienta en Excel para:

1- Registro y consolidación de más de 130 eventos
2- Reducción del tiempo dedicado a consolidar registros.
3- Disminución de errores por captura manual y duplicidad de datos.
4- Mayor trazabilidad y control de la información.
5- Generación más ágil de reportes a las empresas contribuyentes.

<img width="461" height="487" alt="image" src="https://github.com/user-attachments/assets/cab6bb8d-5907-4969-b8e0-ca30af328213" />
<img width="465" height="484" alt="image" src="https://github.com/user-attachments/assets/d1d0fc04-b465-4f8d-9ddf-ae53072c5e88" />

📁 Contenido del repositorio

Sub GuardarFicha()
    
    Dim shFicha As Worksheet
    Dim shBase As Worksheet
    Dim fila As Long
    Dim Asistencia As String
    Dim Evaluacion As String
    
    On Error GoTo ErrorHandler
    
    Set shFicha = Sheets("02_Ficha_Eventos")
    Set shBase = Sheets("01_BD_Eventos")
    
    fila = shBase.Cells(shBase.Rows.Count, "A").End(xlUp).Row + 1
    
    shBase.Rows(fila).ClearContents
    
    ' Guardar datos
    With shBase
        .Cells(fila, "A").Value = shFicha.Range("B9").Value
        .Cells(fila, "B").Value = GetTextFromShape(shFicha, "txtObjetivos")
        .Cells(fila, "C").Value = GetTextFromShape(shFicha, "txtMetodologia")
        .Cells(fila, "D").Value = GetTextFromShape(shFicha, "txtResultados")
        .Cells(fila, "E").Value = GetTextFromShape(shFicha, "txtLecciones")
        .Cells(fila, "F").Value = shFicha.Range("B7").Value
        .Cells(fila, "G").Value = shFicha.Range("B14").Value
        .Cells(fila, "H").Value = shFicha.Range("B15").Value
        .Cells(fila, "I").Value = shFicha.Range("B8").Value
        .Cells(fila, "J").Value = shFicha.Range("B6").Value
        .Cells(fila, "K").Value = shFicha.OLEObjects("TextBox1").Object.Text
        
        If shFicha.Range("C16").Value = True Then
            Asistencia = "Listado original físico"
        ElseIf shFicha.Range("D16").Value = True Then
            Asistencia = "Listado original digital"
        Else
            Asistencia = ""
        End If
        .Cells(fila, "L").Value = Asistencia
        
        .Cells(fila, "M").Value = shFicha.Range("B17").Value
        .Cells(fila, "O").Value = shFicha.Range("B20").Value
        .Cells(fila, "P").Value = shFicha.Range("B18").Value
        .Cells(fila, "Q").Value = shFicha.Range("B19").Value
        
        If shFicha.Range("C21").Value = True Then
            Evaluacion = "Yes"
        ElseIf shFicha.Range("D21").Value = True Then
            Evaluacion = "No"
        Else
            Evaluacion = ""
        End If
        .Cells(fila, "R").Value = Evaluacion
        
        .Cells(fila, "T").Formula = shFicha.Range("B26").Value
        .Cells(fila, "U").Formula = shFicha.Range("B27").Value
        .Cells(fila, "V").Formula = shFicha.Range("B28").Value
        .Cells(fila, "S").Formula = shFicha.Range("B29").Value
    End With
    
    MsgBox "La ficha ha sido guardada correctamente en la base de datos.", vbInformation, "Guardado exitoso"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al guardar la ficha: " & Err.Description, vbCritical, "Error"
    
End Sub

Private Function GetTextFromShape(ws As Worksheet, shapeName As String) As String
    On Error Resume Next
    GetTextFromShape = ws.Shapes(shapeName).TextFrame2.TextRange.Text
    If Err.Number <> 0 Then GetTextFromShape = ""
    On Error GoTo 0
End Function

Sub LimpiarFicha()

    Dim shFicha As Worksheet
    Set shFicha = Sheets("02_Ficha_Eventos")

    'Limpiar celdas
    shFicha.Range("A4:B5").ClearContents
    shFicha.Range("B6").MergeArea.ClearContents
    shFicha.Range("B7").MergeArea.ClearContents
    shFicha.Range("B8").MergeArea.ClearContents
    shFicha.Range("B9").MergeArea.ClearContents

    shFicha.Range("B14").ClearContents
    shFicha.Range("B15").ClearContents
    shFicha.Range("B17").ClearContents
    shFicha.Range("B18").ClearContents
    shFicha.Range("B19").ClearContents
    shFicha.Range("B20").ClearContents
    shFicha.Range("B26").ClearContents
    shFicha.Range("B27").ClearContents
    shFicha.Range("B28").ClearContents
    shFicha.Range("B29").ClearContents
    
    Dim shp As Shape
    
    
    ' Limpiar checkbox de Formulario
    For Each shp In shFicha.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlCheckBox Then
                shp.ControlFormat.Value = xlOff
            End If
        End If
    Next shp
    
    With shFicha.OLEObjects("TextBox1").Object
        .Text = "Escribir el nombre del evento aquí"
        .ForeColor = RGB(0, 0, 0)
    End With
    
    'Limpiar cuadros de texto (shapes)
    shFicha.Shapes("txtObjetivos").TextFrame2.TextRange.Text = ""
    shFicha.Shapes("txtMetodologia").TextFrame2.TextRange.Text = ""
    shFicha.Shapes("txtResultados").TextFrame2.TextRange.Text = ""
    shFicha.Shapes("txtLecciones").TextFrame2.TextRange.Text = ""

    MsgBox "Ficha limpia.", vbInformation

End Sub







