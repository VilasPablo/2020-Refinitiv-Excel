
'
 Sub Indices_Series()
' Macro1 Macro



'
 
Dim ref1 As String
Dim extraccion1 As String
Dim extraccion2 As String
Dim tota1 As Long 'Datos existentes en la REQUEST_TABLE.
Dim datos_analizados As Long 'Por donde vamos en la REQUEST_TABLE
Dim anchura As Integer
Dim fecha_numerica As Long
Dim index_name As String
Dim contador1 As Long
Dim index_month As Long
Dim muestra As Variant
Dim hoja As String
Dim columnas As Long
Dim colum_letra As String
Dim filas As Long
Dim ruta
Dim rng As Range
Dim filas2 As Long
Dim indicador As Long 'Columna1 para pasar rápido de mes en mes

   contador1 = 1
   datos_analizados = 0
   indicador = 10000
    ruta = ThisWorkbook.Path
    Sheets("REQUEST_TABLE").Select
    tota1 = Range("B7").CurrentRegion.Rows.Count 'Todos los Datos descargados en la REQUEST_TABLE

    For tt = datos_analizados To tota1 - 1
    ref1 = Cells(datos_analizados + 6 + 1, 5)
    anchura = Len(ref1)
    extraccion1 = Mid(ref1, 1, anchura - 4) 'Nombre del índice
    filas2 = 0
    ' datos mensuales del índice, hubiera sido más eficiente usar la función contar
        Do Until contador1 > tota1
        ref1 = Cells(contador1 + 6, 5)
        anchura = Len(ref1)
        extraccion2 = Mid(ref1, 1, anchura - 4)
        If extraccion1 = extraccion2 Then
        index_month = index_month + 1
        Else
        End If
        contador1 = contador1 + 1
        Loop
        contador1 = 1 'reiniciamos contador
    'Crear un excel con el nombre del índice
    ref1 = Cells(datos_analizados + 6 + 1, 5)
    anchura = Len(ref1)
    extraccion1 = Mid(ref1, 1, anchura - 4) 'ticker del índice
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=ruta & "\" & extraccion1
    Sheets("Hoja1").Name = extraccion1
    ThisWorkbook.Activate
        
    'Ir hoja por hoja extrayendo los datos del índice
    For i = datos_analizados To index_month - 1
    ref1 = Cells(datos_analizados + 6 + 1, 5)
    anchura = Len(ref1)
    extraccion1 = Mid(ref1, 1, anchura - 4)
    index_name = extraccion1
    'Sacar del ticker de descarga mensual LA FECHA
    extraccion1 = Mid(ref1, anchura - 4 + 1, 2)
    extraccion2 = Mid(ref1, anchura - 2 + 1, 2)
    If extraccion2 > 30 Then
    extraccion2 = 19 & extraccion2
    fecha_numerica = extraccion2 & extraccion1
    Else
    extraccion2 = 20 & extraccion2
    fecha_numerica = extraccion2 & extraccion1
    End If
    'Número de hoja donde se encuentran los datos
    ref1 = Cells(datos_analizados + 6 + 1, 11)
    anchura = Len(ref1)
    extraccion1 = Mid(ref1, 3, anchura - 8)
    hoja = extraccion1
    datos_analizados = datos_analizados + 1
    'Copiar los datos de la descarga mensual. Filas = Datos de la descarga mensual. Filas2= Número de datos pasado al nueveo Excel
    filas = Sheets("" & hoja & "").Range("A1").CurrentRegion.Rows.Count - 1
    columnas = Sheets("" & hoja & "").Range("A1").CurrentRegion.Columns.Count
    colum_letra = Mid(Split(Columns(columnas).Address, ":")(1), 2)

    If filas = 0 Then
        filas = 1
        columnas = Workbooks("" & index_name & "").Sheets("" & index_name & "").Range("A1").CurrentRegion.Columns.Count - 3
        For kk = 2 To columnas
        Sheets("" & hoja & "").Cells(1, kk).Value = "ERROR SERVIDOR"
        Next
        Worksheets("" & hoja & "").Range(Worksheets("" & hoja & "").Cells(1, 1), _
        Worksheets("" & hoja & "").Cells(filas + 1, columnas)).Copy
    Else
    Worksheets("" & hoja & "").Range(Worksheets("" & hoja & "").Cells(2, 1), _
    Worksheets("" & hoja & "").Cells(filas + 1, columnas)).Copy
    End If
    

    'Definir rango de la hoja creada y pegar los datos
    Set rng = Workbooks("" & index_name & "").Sheets("" & index_name & "").Range(Workbooks("" & index_name & "").Worksheets("" _
    & index_name & "").Cells(2 + filas2, 2), _
    Workbooks("" & index_name & "").Worksheets("" & index_name & "").Cells(filas2 + filas + 1, columnas + 1))
    ActiveSheet.Paste Destination:=rng
    Workbooks("" & index_name & "").Sheets("" & index_name & "").Range(Workbooks("" & index_name & "").Worksheets("" _
    & index_name & "").Cells(2 + filas2, 1), Workbooks("" & index_name & "").Worksheets("" _
    & index_name & "").Cells(2 + filas2, 1)).Value = indicador '10.000
    Workbooks("" & index_name & "").Sheets("" & index_name & "").Range(Workbooks("" & index_name & "").Worksheets("" _
    & index_name & "").Cells(2 + filas2, columnas + 2), _
    Workbooks("" & index_name & "").Worksheets("" & index_name & "").Cells(filas2 + filas + 1, columnas + 2)).Value = fecha_numerica 'Columna con la fecha
    Workbooks("" & index_name & "").Sheets("" & index_name & "").Range(Workbooks("" & index_name & "").Worksheets("" _
    & index_name & "").Cells(2 + filas2, columnas + 3), _
    Workbooks("" & index_name & "").Worksheets("" & index_name & "").Cells(filas2 + filas + 1, columnas + 3)).Value = hoja ' En que hoja estan los datos.
 
    filas2 = filas2 + filas 'Contador para pegar los datos en forma de listado uno detrás de otro.
    indicador = indicador + 1
    
    Next
    'Pegar títulos de los datos descargados
    Worksheets("" & hoja & "").Range(Worksheets("" & hoja & "").Cells(1, 1), _
    Worksheets("" & hoja & "").Cells(1, columnas)).Copy
    Set rng = Workbooks("" & index_name & "").Sheets("" & index_name & "").Range(Workbooks("" & index_name & "").Worksheets("" _
    & index_name & "").Cells(1, 2), _
    Workbooks("" & index_name & "").Worksheets("" & index_name & "").Cells(1, columnas + 1))
    ActiveSheet.Paste Destination:=rng
    Workbooks("" & index_name & "").Worksheets("" & index_name & "").Cells(1, columnas + 2).value = "Fecha Numerica"
    Workbooks("" & index_name & "").Worksheets("" & index_name & "").Cells(1, columnas + 3).value = "Ubicación"
    Workbooks("" & index_name & "").Worksheets("" & index_name & "").Cells(1, 1).Value = "Sig. Mes"

    'Eliminar NA que se encuentran en las primeras celdas
    Do While Workbooks("" & index_name & "").Worksheets("" & index_name & "").Cells(2, 4) = "NA" and _
    Workbooks("" & index_name & "").Worksheets("" & index_name & "").Cells(2, 3) = "NA"
        Workbooks("" & index_name & "").Worksheets("" & index_name & "").Rows("2:2").Delete Shift:=xlUp
    Loop
  
 
    Workbooks("" & index_name & "").Close (True)
    Application.Wait (Now + TimeValue("00:00:30"))
    tt = datos_analizados
    
    Next
    
   MsgBox "Los archivos se han guardado en: " & ruta & "", , "Cualquier sugerencia escriba a: pablo.vilas.naval@gmail.com"
    
End Sub
