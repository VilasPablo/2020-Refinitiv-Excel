Public cont_list As Integer
Public Nº_atributos As Integer


Sub Creación_listado_para_Time_Series()
'
' Macro1 Macro
'

'
Dim Nº_ISINs As Integer
Dim TOTAL As Integer 'Total descargas
Dim hoja As String
Dim ISIN_name As String 'Un ISIN en concreto
Dim contador1 As Integer 'Contador para asignar ISIns al array
Dim salto_título As Long '"dígito de control"
Dim controlador_listado As Long 'colocar bien la info en forma de listado

Dim nºfilas As Variant
Dim FILAS_HOJA As Long
Dim DEAD As Long   'celda de muerte en "hoja"
Dim BIRTH As Long  'celda nacimiento en "hoja"
Dim INICIO As Long 'comienzo de los datos ISIN para listado
Dim FINAL As Long  'final de los datos ISIN para listado
Dim contador_while As Long 'nos ayuda a asignar los valores de DEAD y BIRTH
Dim filas_por_listado As Long
Dim Número_empresas_time As Long

Call Ordenar

filas_por_listado = Application.InputBox("Introduzca el número máximo de filas que desea por listado(únicamente introduzca números)", _
"El máximo de filas permitido por Excel es 1.048.576")
If IsNumeric(filas_por_listado) = False Or filas_por_listado < 2 Then
filas_por_listado = 1000000
End If
controlador_listado = 2
salto_título = 10001
Número_empresas_time = Worksheets("TIME-BTIME").Range("A1").CurrentRegion.Rows.Count
Worksheets.Add.Name = "Listado"

'TOTAL de hojas descargadas
Do While Worksheets("REQUEST_TABLE").Cells(7 + Z1, 5) <> ""
    Z1 = Z1 + 1
    TOTAL = TOTAL + 1
    Loop
TOTAL = TOTAL + 6

N_filas_time = Worksheets("TIME-BTIME").Range("A1").CurrentRegion.Rows.Count
For fil = 1 To N_filas_time
    Worksheets("TIME-BTIME").Cells(fil, 5).Value = fil
Next

'Bucle para ir yendo por todas las hojas
For RESQUEST = 7 To TOTAL
    Nº_atributos = 0
    Nº_ISINs = 0
    
    'Averiguamos el total de atributos (es importante que después de cada dattype exista coma)
    lenn = Len(Worksheets("REQUEST_TABLE").Cells(RESQUEST, 6))
    For X2 = 1 To lenn
        If Mid(Worksheets("REQUEST_TABLE").Cells(RESQUEST, 6), X2, 1) = "," Then
            Nº_atributos = 1 + Nº_atributos
        End If
    Next
    
    'Averiguamos el total de ISINs (después de cada ISIN debe haber coma)
    lenn = Len(Worksheets("REQUEST_TABLE").Cells(RESQUEST, 5))
    For X1 = 1 To lenn
        If Mid(Worksheets("REQUEST_TABLE").Cells(RESQUEST, 5), X1, 1) = "," Then
            Nº_ISINs = 1 + Nº_ISINs
        End If
    Next
    'Matriz almacenadora de ISINs, BTIME, TIME
    ReDim matriz(Nº_ISINs - 1, 2)
    contador1 = 0
    'Asiganmos a la matriz los valores
    For X11 = 1 To lenn
        If Mid(Worksheets("REQUEST_TABLE").Cells(RESQUEST, 5), X11, 1) = "," Then
            matriz(contador1, 0) = ISIN_name
            If Número_empresas_time < 2 Then GoTo timee
            nºfilas = Application.WorksheetFunction.VLookup(ISIN_name, Range(Worksheets("TIME-BTIME").Cells(2, 1), _
            Worksheets("TIME-BTIME").Cells(N_filas_time, 5)), 5, True)
                If Worksheets("TIME-BTIME").Cells(nºfilas, 1) = ISIN_name Then
                    If Len(Worksheets("TIME-BTIME").Cells(nºfilas, 2)) = 10 Then
                        matriz(contador1, 1) = Mid(Worksheets("TIME-BTIME").Cells(nºfilas, 2), 7, 4) & Mid(Worksheets("TIME-BTIME") _
                        .Cells(nºfilas, 2), 4, 2) & Mid(Worksheets("TIME-BTIME").Cells(nºfilas, 2), 1, 2)
                    Else
timee:
                        matriz(contador1, 1) = 100000000
                    End If
                    If Número_empresas_time < 2 Then GoTo btimee
                    If Len(Worksheets("TIME-BTIME").Cells(nºfilas, 3)) = 10 Then
                        matriz(contador1, 2) = Mid(Worksheets("TIME-BTIME").Cells(nºfilas, 3), 7, 4) & Mid(Worksheets("TIME-BTIME") _
                        .Cells(nºfilas, 3), 4, 2) & Mid(Worksheets("TIME-BTIME").Cells(nºfilas, 3), 1, 2)
                    Else
btimee:
                        matriz(contador1, 2) = 0
                    End If
                End If
            ISIN_name = ""
            contador1 = contador1 + 1
        Else
            If Mid(Worksheets("REQUEST_TABLE").Cells(RESQUEST, 5), X11, 1) <> " " Then
                ISIN_name = ISIN_name + Mid(Worksheets("REQUEST_TABLE").Cells(RESQUEST, 5), X11, 1)
            End If
        End If
    Next
    
    'hoja donde extraemos datos para listados
    hoja = Mid(Worksheets("REQUEST_TABLE").Cells(RESQUEST, 11), 3, Len(Worksheets("REQUEST_TABLE").Cells(RESQUEST, 11)) - 8)
    'total de filas en "hoja"
    FILAS_HOJA = Worksheets("" & hoja & "").Range("A1").CurrentRegion.Rows.Count
        
    'Desntro de la misma hoja vamos título por título
    For xx1 = 1 To Nº_ISINs
        'reinicios de valores
        FINAL = FILAS_HOJA
        INICIO = 4
        contador_while = 4
        BIRTH = 0
        DEAD = 0
        
        ' si el título ha nacido después de la última fecha descargada pasamos al siguiente título
        If IsEmpty(Worksheets("" & hoja & "").Cells(4, 1)) = True Then GoTo lineMM
        If CLng(matriz(xx1 - 1, 2)) > CLng(Mid(Worksheets("" & hoja & "").Cells(FINAL, 1), 7, 4) & _
        Mid(Worksheets("" & hoja & "").Cells(FINAL, 1), 4, 2) & Mid(Worksheets("" & hoja & "").Cells(FINAL, 1), 1, 2)) Then GoTo lineff
        
        'si el títutlo ha muerto antes que la primera fecha descargada pasamos al siguiente título
        If IsEmpty(Worksheets("" & hoja & "").Cells(4, 1)) = True Then GoTo lineMM
        If CLng(matriz(xx1 - 1, 1)) < CLng(Mid(Worksheets("" & hoja & "").Cells(INICIO, 1), 7, 4) & _
        Mid(Worksheets("" & hoja & "").Cells(INICIO, 1), 4, 2) & Mid(Worksheets("" & hoja & "").Cells(INICIO, 1), 1, 2)) Then GoTo lineff
        
        'Diferencia entra nacimiento y primera fecha descargada
        Do While CLng(matriz(xx1 - 1, 2)) > CLng(Mid(Worksheets("" & hoja & "").Cells(contador_while, 1), 7, 4) & Mid(Worksheets _
            ("" & hoja & "").Cells(contador_while, 1), 4, 2) & Mid(Worksheets("" & hoja & "").Cells(contador_while, 1), 1, 2)) _
            And contador_while < FINAL
            BIRTH = BIRTH + 1
            contador_while = contador_while + 1
        Loop
        contador_while = FINAL - 1
        
        'Diferencia entre muerte y última fecha descargada
        Do While CLng(matriz(xx1 - 1, 1)) <= CLng(Mid(Worksheets("" & hoja & "").Cells(contador_while, 1), 7, 4) & Mid(Worksheets _
            ("" & hoja & "").Cells(contador_while, 1), 4, 2) & Mid(Worksheets("" & hoja & "").Cells(contador_while, 1), 1, 2)) _
            And contador_while > INICIO
            DEAD = DEAD + 1
            contador_while = contador_while - 1
        Loop
lineMM:
        'Límite inferior y superior para llevar a listado
        INICIO = INICIO + BIRTH
        FINAL = FINAL - DEAD
        'Copias y pegamos en listado la información relativa a los DATATYPES(Atributos)
        Range(Worksheets("" & hoja & "").Cells(INICIO, 2 + Nº_atributos * (xx1 - 1)), _
        Worksheets("" & hoja & "").Cells(FINAL, 1 + Nº_atributos * xx1)).Copy
        ActiveSheet.Paste Destination:=Worksheets("Listado").Range("F" & controlador_listado)
        'Copiamos pegamos fechas
        Range(Worksheets("" & hoja & "").Cells(INICIO, 1), Worksheets("" & hoja & "").Cells(FINAL, 1)).Copy
        ActiveSheet.Paste Destination:=Worksheets("Listado").Range("B" & controlador_listado)
        'Dejamos indicado en que hoja esta la información de este título
        Range(Worksheets("Listado").Cells(controlador_listado, 6 + Nº_atributos), Worksheets("Listado").Cells(controlador_listado _
        + FINAL - INICIO, 6 + Nº_atributos)).Value = hoja
        'Asiganmos el salto de título /dígito que delimita títulos
        Worksheets("Listado").Range("A" & controlador_listado).Value = salto_título
    
        'Borrar cuando Laura de el visto bueno a borrar según BTIME y TIME
        'If BIRTH <> 0 Then
        'tata = Application.WorksheetFunction.Index(matriz, xx1, 3)
        'Worksheets("Listado").Cells(controlador_listado + BIRTH, 22).Value = CStr("'" & Mid(tata, 7, 2) & "/" & Mid(tata, 5, 2) & "/" & Mid(tata, 1, 4))
        'End If
        'If DEAD <> 0 Then
        'tata = Application.WorksheetFunction.Index(matriz, xx1, 2)
        'Worksheets("Listado").Cells(FINAL - INICIO + controlador_listado - DEAD + 1, 23).Value = CStr("'" & Mid(tata, 7, 2) & "/" & Mid(tata, 5, 2) & "/" _
        '& Mid(tata, 1, 4))
        'End If
        
        'Chequeamos que al menos un ISIN de la columna 2 coincida con el de la request_table
        For xtt = 0 To Nº_atributos - 1
            If matriz(xx1 - 1, 0) = Mid(Sheets("" & hoja & "").Cells(2, 2 + Nº_atributos * (xx1 - 1) + xtt), _
                1, Len(matriz(xx1 - 1, 0))) Then
                Exit For
            Else
                If xtt = Nº_atributos - 1 Then
                'Pintar de rojo las columnas C,D,E cuando el ISIN de la Resquest table no coincide con ninguno de los
                'descargados en la columna 2, en la columna E se pondrá el isin de la request table
                    Worksheets("Listado").Range("C" & controlador_listado & ":E" & FINAL - INICIO + controlador_listado).Interior.ColorIndex = 3
                    Worksheets("Listado").Range("E" & controlador_listado & ":E" & FINAL - INICIO + controlador_listado).Value = _
                    Application.WorksheetFunction.Index(matriz, xx1, 1)
                    GoTo line22
                End If
            End If
        Next
        
        'Copiamos y pegamos NOMBRE, ISIN y CURREN
        Range(Worksheets("" & hoja & "").Cells(1, 2 + Nº_atributos * (xx1 - 1) + xtt), _
        Worksheets("" & hoja & "").Cells(3, 2 + Nº_atributos * (xx1 - 1) + xtt)).Copy
        Worksheets("Listado").Range("C" & controlador_listado & ":E" & FINAL - INICIO + _
        controlador_listado).PasteSpecial Transpose:=True 'NOMBRE, ISIN CURREN
        'Ponemos el código (ISIN, TICKER) sin el código del DATATYPE
        Worksheets("Listado").Range("D" & controlador_listado & ":D" & FINAL - INICIO + _
        controlador_listado).Value = Application.WorksheetFunction.Index(matriz, xx1, 1)
line22:
        'Ajustamos variables relacionadas con el listado para los siguientes títulos
        salto_título = salto_título + 1
        controlador_listado = FINAL - INICIO + 1 + controlador_listado
        
        'Si el listado ha superado cierto número de filas, se guarda y se crea uno nuevo
        If controlador_listado > filas_por_listado Then
            Application.Wait (Now + TimeValue("00:00:05"))
            Call guardar_list
            controlador_listado = 2
            Worksheets.Add.Name = "Listado"
        End If
lineff:
    Next
           
Next
Application.Wait (Now + TimeValue("00:00:05"))
Call guardar_list
Worksheets("TIME-BTIME").Columns(5).ClearContents
MsgBox "Los archivos se han guardado en:" & CStr(ThisWorkbook.Path), , "Cualquier sugerencia escriba a: pablo.vilas.naval@gmail.com"

End Sub

Sub Asignar_títulos()

'Asignamos los rótulos a los listados SEGÚN DATATYPES(MV, MVC, P...)
Dim y
ReDim y(Nº_atributos - 1, 1)
lenn = Len(Worksheets("REQUEST_TABLE").Cells(7, 6))
contador1 = 0
For X11 = 1 To lenn
    If Mid(Worksheets("REQUEST_TABLE").Cells(7, 6), X11, 1) = "," Then
        y(contador1, 1) = ISIN_name
        ISIN_name = ""
        contador1 = contador1 + 1
    Else
        If Mid(Worksheets("REQUEST_TABLE").Cells(7, 6), X11, 1) <> " " Then
            ISIN_name = ISIN_name + Mid(Worksheets("REQUEST_TABLE").Cells(7, 6), X11, 1)
        End If
    End If
Next

For X3 = 1 To Nº_atributos
    Worksheets("Listado").Cells(1, 5 + X3).Value = Application.WorksheetFunction.Index(y, X3, 2)
Next

Worksheets("Listado").Cells(1, 6 + Nº_atributos).Value = "Nº Hoja"
Worksheets("Listado").Cells(1, 1).Value = "Cmb. Serie"
Worksheets("Listado").Cells(1, 2).Value = "Date"
Worksheets("Listado").Cells(1, 3).Value = "Name"
Worksheets("Listado").Cells(1, 4).Value = "Code"
Worksheets("Listado").Cells(1, 5).Value = "Currency"

End Sub


Sub guardar_list()
Call Asignar_títulos
Dim ruta As String

cont_list = cont_list + 1
ruta = ThisWorkbook.Path
Workbooks.Add
Application.Wait (Now + TimeValue("00:00:05"))
ActiveSheet.Name = "Listado" & cont_list
ThisWorkbook.Worksheets("Listado").Cells.Copy
Application.Wait (Now + TimeValue("00:00:05"))
ActiveSheet.Paste Destination:=Range("A1")
Application.Wait (Now + TimeValue("00:00:05"))
    Call dejar_bonito
ActiveWorkbook.SaveAs ruta & "\Listado" & cont_list
Set wbk = ActiveWorkbook
ThisWorkbook.Activate
wbk.Close (False)
Application.DisplayAlerts = False
Sheets("Listado").Delete
Application.DisplayAlerts = True
End Sub




Sub dejar_bonito()

    ActiveSheet.Cells.ColumnWidth = 11
    ActiveSheet.Rows("1:1").RowHeight = 45
    
    With ActiveSheet.Cells
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With ActiveSheet.Rows("1:1")
        .Font.Bold = True
    End With
    ActiveSheet.Rows("2:" & Range("B1").CurrentRegion.Rows.Count).RowHeight = 13.5
End Sub

Sub Ordenar()
'
Dim Columna_ordenado As Integer
Dim letra As String

filitas = Sheets("Time-BTIME").Range("A1").CurrentRegion.Rows.Count
For ttt = 2 To filitas
Sheets("Time-BTIME").Cells(ttt, 1).Value = "'" & Sheets("Time-BTIME").Cells(ttt, 1)
Next

Columna_ordenado = 1 ' ordenar según esta columna
letra = Split(Cells(1, Columna_ordenado).Address, "$")(1)

ActiveWorkbook.Worksheets("TIME-BTIME").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("TIME-BTIME").Sort.SortFields.Add Key:=Range _
        ("" & letra & "2:" & letra & Range(letra & "1").CurrentRegion.Rows.Count), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("TIME-BTIME").Sort
        .SetRange Range("A1:ZL" & Range("C1").CurrentRegion.Rows.Count)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub








