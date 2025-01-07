' Procedimiento Principal: Copiar Datos de Retiro de Pollo
Sub CopiarDatosDeRetiroPollo()
    Dim folderPath As String
    Dim sourceFileName As String
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim destinoWorkbook As Workbook
    Dim destinoSheet As Worksheet
    Dim tablaDestino As ListObject
    Dim ultimaFila As Long
    Dim rutaArchivo As String
    Dim fila As Long
    Dim numFilas As Long
    Dim destinoSeleccionada As String
    
    ' 1. Definir la carpeta donde se encuentran los archivos
    folderPath = "D:\Users\lcontreras\Desktop\YOUR_FILES\"
    
    ' 2. Definir el nombre del archivo fuente
    sourceFileName = "NAME_FILE.xlsx"
    
    ' 3. Definir la ruta completa del archivo fuente
    rutaArchivo = folderPath & sourceFileName
    
    ' 4. Verificar que la carpeta exista
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "La carpeta especificada no existe: " & folderPath, vbCritical
        LogEvent "Error: Carpeta no encontrada: " & folderPath
        Exit Sub
    End If
    
    ' 5. Verificar que el archivo fuente exista
    If Dir(rutaArchivo) = "" Then
        MsgBox "El archivo fuente no se encontró: " & rutaArchivo, vbCritical
        LogEvent "Error: Archivo fuente no encontrado: " & rutaArchivo
        Exit Sub
    End If
    
    LogEvent "Inicio de la macro. Archivo fuente: " & sourceFileName
    
    ' 6. Abrir el archivo fuente
    On Error Resume Next
    Set sourceWorkbook = Workbooks.Open("D:\Users\lcontreras\Desktop\Excel diarios de OUTLOOK\NAME_FILE.xlsx")
    If sourceWorkbook Is Nothing Then
        MsgBox "No se pudo abrir el archivo fuente: " & rutaArchivo, vbCritical
        LogEvent "Error: No se pudo abrir el archivo fuente: " & rutaArchivo
        Exit Sub
    End If
    On Error GoTo 0
    
    LogEvent "Archivo fuente abierto correctamente: " & rutaArchivo
    
    ' 7. Definir la hoja de origen
    On Error Resume Next
    Set sourceSheet = sourceWorkbook.Sheets("Programa FAE ")
    On Error GoTo 0
    
    If sourceSheet Is Nothing Then
        MsgBox "La hoja 'Programa FAE ' no existe en el archivo fuente.", vbCritical
        LogEvent "Error: La hoja 'Programa FAE ' no existe en el archivo fuente."
        sourceWorkbook.Close SaveChanges:=False
        Exit Sub
    End If
    
    LogEvent "Hoja de origen encontrada: Programa FAE "
    
    ' 8. Definir el libro de destino (el que contiene esta macro)
    Set destinoWorkbook = ThisWorkbook
    
    ' 9. Seleccionar la hoja de destino
    destinoSeleccionada = InputBox("Ingrese el nombre de la hoja de destino donde se pegarán los datos (por ejemplo, ENE.25):", "Seleccionar Hoja de Destino", "ENE.25")
    
    If Trim(destinoSeleccionada) = "" Then
        MsgBox "No se ingresó ningún nombre de hoja de destino. La macro se detendrá.", vbExclamation
        LogEvent "Error: No se ingresó ningún nombre de hoja de destino."
        sourceWorkbook.Close SaveChanges:=False
        Exit Sub
    End If
    
    On Error Resume Next
    Set destinoSheet = destinoWorkbook.Sheets(destinoSeleccionada)
    On Error GoTo 0
    
    If destinoSheet Is Nothing Then
        MsgBox "La hoja de destino '" & destinoSeleccionada & "' no existe en el libro '" & destinoWorkbook.Name & "'.", vbCritical
        LogEvent "Error: La hoja de destino '" & destinoSeleccionada & "' no existe en el libro '" & destinoWorkbook.Name & "'."
        sourceWorkbook.Close SaveChanges:=False
        Exit Sub
    End If
    
    LogEvent "Hoja de destino seleccionada: " & destinoSeleccionada
    
    ' 10. Definir la tabla de destino
    On Error Resume Next
    ' Reemplaza "BaseTransporte.Sept2315161724578910111213141516" para cada mes
    Set tablaDestino = destinoSheet.ListObjects("BaseTransporte.Sept2315161724578910111213141516")
    On Error GoTo 0
    
    If tablaDestino Is Nothing Then
        MsgBox "La tabla no existe en la hoja '" & destinoSeleccionada & "'.", vbCritical
        LogEvent "Error: La tabla no existe en la hoja '" & destinoSeleccionada & "'."
        sourceWorkbook.Close SaveChanges:=False
        Exit Sub
    End If
    
    LogEvent "Tabla de destino encontrada"
    
    ' 11. Encontrar la última fila con datos en la hoja de origen, empezando desde E10
    With sourceSheet
        numFilas = .Cells(.Rows.Count, "E").End(xlUp).Row
        If numFilas < 10 Then
            MsgBox "No hay datos para copiar en la hoja 'Programa FAE ' a partir de la fila 10.", vbInformation
            LogEvent "Información: No hay datos para copiar a partir de la fila 10."
            sourceWorkbook.Close SaveChanges:=False
            Exit Sub
        End If
    End With
    
    LogEvent "Número de filas con datos a procesar: " & (numFilas - 9)
    
    ' 12. Iterar desde la fila 10 hasta la última fila con datos
    For fila = 10 To numFilas
        ' Verificar si la fila tiene datos en la columna E
        If Trim(sourceSheet.Cells(fila, "E").Value) <> "" Then
            ' Agregar una nueva fila a la tabla de destino
            tablaDestino.ListRows.Add
            ultimaFila = tablaDestino.ListRows.Count
            
            ' Asignar valores de las columnas E, H, J, K, L, M, N, S del origen a A, B, C, D, E, F, G, H del destino
            With tablaDestino.ListRows(ultimaFila).Range
                .Cells(1, tablaDestino.ListColumns("Frigorifico").Index).Value = sourceSheet.Cells(fila, "E").Value ' Columna E ? A
                .Cells(1, tablaDestino.ListColumns("Granja").Index).Value = sourceSheet.Cells(fila, "H").Value ' Columna H ? B
                .Cells(1, tablaDestino.ListColumns("Galpón").Index).Value = sourceSheet.Cells(fila, "I").Value ' Columna J ? C
                .Cells(1, tablaDestino.ListColumns("Tipo").Index).Value = sourceSheet.Cells(fila, "J").Value ' Columna K ? D
                .Cells(1, tablaDestino.ListColumns("Sexo").Index).Value = sourceSheet.Cells(fila, "K").Value ' Columna L ? E
                .Cells(1, tablaDestino.ListColumns("Edad").Index).Value = sourceSheet.Cells(fila, "L").Value ' Columna M ? F
                .Cells(1, tablaDestino.ListColumns("Peso Promedio").Index).Value = sourceSheet.Cells(fila, "M").Value ' Columna N ? G
                .Cells(1, tablaDestino.ListColumns("Cantidad").Index).Value = sourceSheet.Cells(fila, "N").Value ' Columna S ? H
                
                ' Asignar valores de las columnas AC, AD, AE, AF del origen a L, M, N, O del destino
                .Cells(1, tablaDestino.ListColumns("Llegada a Matadero").Index).Value = sourceSheet.Cells(fila, "AC").Value ' Columna AC ? L
                .Cells(1, tablaDestino.ListColumns("Inicio de Descarguío").Index).Value = sourceSheet.Cells(fila, "AD").Value ' Columna AD ? M
                .Cells(1, tablaDestino.ListColumns("Inicio del Colgado").Index).Value = sourceSheet.Cells(fila, "AE").Value ' Columna AE ? N
                .Cells(1, tablaDestino.ListColumns("Finalización Descarguío").Index).Value = sourceSheet.Cells(fila, "AF").Value ' Columna AF ? O
            End With
            
            LogEvent "Datos copiados de la fila " & fila & " al destino."
        End If
    Next fila
    
    ' 13. Cerrar el archivo de origen sin guardar cambios
    sourceWorkbook.Close SaveChanges:=False
    LogEvent "Archivo fuente cerrado sin guardar cambios: " & rutaArchivo
    
    ' 14. Mensaje de confirmación
    MsgBox "Los datos han sido copiados correctamente desde '" & sourceFileName & "' a la hoja '" & destinoSeleccionada & "'.", vbInformation
    LogEvent "Proceso completado exitosamente."
End Sub

' Procedimiento de Registro de Eventos (Logging) Interno
Sub LogEvent(message As String)
    Dim destinoWorkbook As Workbook
    Dim logSheet As Worksheet
    Dim ultimaFila As Long
    
    ' Definir el libro de destino (el que contiene esta macro)
    Set destinoWorkbook = ThisWorkbook
    
    ' Definir la hoja de log
    On Error Resume Next
    Set logSheet = destinoWorkbook.Sheets("Log")
    On Error GoTo 0
    
    ' Verificar que la hoja de log exista
    If logSheet Is Nothing Then
        ' Crear la hoja de log si no existe
        Set logSheet = destinoWorkbook.Sheets.Add(After:=destinoWorkbook.Sheets(destinoWorkbook.Sheets.Count))
        logSheet.Name = "Log"
        ' Configurar encabezados
        logSheet.Range("A1").Value = "Fecha y Hora"
        logSheet.Range("B1").Value = "Mensaje"
    End If
    
    ' Encontrar la última fila con datos
    ultimaFila = logSheet.Cells(logSheet.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Escribir el mensaje con marca de tiempo
    logSheet.Cells(ultimaFila, "A").Value = Now
    logSheet.Cells(ultimaFila, "B").Value = message
End Sub
