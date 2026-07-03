#Tablero interactivo para la validaciòn de soportes técnicos

🎯 ¿Qué problema resuelve?

La dependencia de procesos manuales para validar y realizar el seguimiento de más de 75.000 soportes técnicos en formato PDF,
lo que generaba altos tiempos de revisión, riesgo de inconsistencias en la información y dificultades para monitorear oportunamente 
los indicadores del proyecto.

🛠️ Herramientas: Excel | Visual Basic

📈 Resultados: 

Macro en Excel para el seguimiento y validación de más de 75 mil soportes técnicos en formato PDF, 
optimizando el monitoreo de indicadores y control de información del proyecto nacional “Renovación de 5.000 predios cacaoteros”.

📁 Contenido del repositorio

Sub ListarCarpetasYArchivos()
    ' Variables a utilizar en la macro
    Dim CarpetaRaiz As String
    Dim FSO As Object
    Dim Carpeta As Object
    Dim SubCarpeta As Object
    Dim Archivo As Object
    Dim Fila As Integer

    ' Solicita al usuario que ingrese la ruta de la carpeta
    CarpetaRaiz = InputBox("Ingresa la ruta de la carpeta a importar:")

    If CarpetaRaiz = "" Then
        Exit Sub
    ElseIf Right(CarpetaRaiz, 1) <> "\" Then
        CarpetaRaiz = CarpetaRaiz & "\"
    End If

    ' Inicializa la fila en la que se empezarán a listar los nombres de carpetas y archivos
    Fila = 1

    ' Crea un objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' Abre la carpeta raíz
    Set Carpeta = FSO.GetFolder(CarpetaRaiz)

    ' Habilita las actualizaciones de pantalla para mejorar el rendimiento
    Application.ScreenUpdating = True

    ' Listar carpetas y archivos
    ListarCarpetasYArchivosRecursivo Carpeta, Fila

    ' Deshabilita las actualizaciones de pantalla
    Application.ScreenUpdating = False
End Sub

Sub ListarCarpetasYArchivosRecursivo(CarpetaActual As Object, ByRef Fila As Integer)
    Dim Archivo As Object
    Dim SubCarpeta As Object

    ' Listar carpetas (mostrar solo el nombre de la carpeta)
    Cells(Fila, 1).Value = CarpetaActual.Name
    Fila = Fila + 1

    ' Listar archivos en esta carpeta
    For Each Archivo In CarpetaActual.Files
        Cells(Fila, 2).Value = Archivo.Name
        Fila = Fila + 1
    Next Archivo

    ' Recorrer las subcarpetas
    For Each SubCarpeta In CarpetaActual.SubFolders
        ListarCarpetasYArchivosRecursivo SubCarpeta, Fila
    Next SubCarpeta
End Sub

