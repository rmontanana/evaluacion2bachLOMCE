Attribute VB_Name = "Evaluaciones"
Public Const miEvaluacion As Integer = 1
Public Const miRecuperacion As Integer = 2
Public Const miOtras As Integer = 3
Public evaluaciones(3) As String
Public recuperaciones(3) As String
Public otras(2) As String
Sub estableceVersion(ByRef libro As Workbook)
    Dim hoja As Worksheet
    Set hoja = libro.Worksheets("Estructura")
    Call desprotege(hoja)
    hoja.Range("d13") = version
    Call protege(hoja)
End Sub
Sub Fill_Evaluaciones()
    Dim libro As Workbook
    Dim actual As Worksheet
    
    Call inicializa
    Call inicializaHojasEvaluacion(evaluaciones, recuperaciones, otras)
    Set actual = ActiveSheet
    actual.Range("B4").Select
    Set libro = Workbooks.Open(ThisWorkbook.Path & Application.PathSeparator & libroEvaluacion)
    Call estableceVersion(libro)
    contador = 0
    ' Evaluaciones
    For evaluacion = 1 To 3
        nombre = evaluaciones(evaluacion)
        Call procesaEvaluacion(actual, libro, nombre, evaluacion, miEvaluacion)
        contador = contador + 1
    Next
    ' Recuperaciones
    For Recuperacion = 1 To 3
        Call procesaEvaluacion(actual, libro, recuperaciones(Recuperacion), Recuperacion, miRecuperacion)
        contador = contador + 1
    Next
    ' Ordinaria y Extraordinaria
    For otra = 1 To 2
        Call procesaEvaluacion(actual, libro, otras(otra), otra, miOtra)
        contador = contador + 1
    Next
    ' Fin
    actual.Cells(4, 3).Value = "Fin de Proceso."
    MsgBox "Se han procesado " + Str(contador) + " hojas"""
    Call finaliza
    libro.Protect password, True
    Call libro.Save
End Sub
Sub procesaEvaluacion(actual, libro, nombre, evaluacion, tipoHoja)
    Dim hoja As Worksheet
    Dim row As Integer, column As Integer
    Dim rangos(3, 6, 2) As Integer
    
    Call estableceRangosEvaluacion(rangos)
    iniCol = 4
    finCol = 63
    Set hoja = libro.Worksheets(nombre)
    Call desprotege(hoja)
    For rango = 1 To 6
        If tipoHoja = miOtro Then
             ' Las hojas de este tipo tienen todos los rangos como la 3»
            indiceEvaluacion = 3
        Else
            indiceEvaluacion = evaluacion
        End If
        If rangos(indiceEvaluacion, rango, 1) <> 0 Then
            actual.Cells(4, 3).Value = "Evaluaci—n " + nombre + " Rango:" + Str(rango)
            DoEvents
            For column = iniCol To finCol Step 2
                iniRango = rangos(indiceEvaluacion, rango, 1)
                finRango = rangos(indiceEvaluacion, rango, 2)
                For row = iniRango To finRango
                    If row = iniRango Then
                        ' Genera la f—rmula para el c‡lculo de la calificaci—n del bloque de contenidos
                        tope = Trim(Str(finRango))
                        Call calificacionBloque(hoja, row, column, tope)
                    End If
                    If column = iniCol Then
                        ' Genera el peso del criterio de calificaci—n
                        Call pesoCriterio(hoja, row, column, rango, evaluacion, tipoHoja)
                    End If
                    Call calificacionCriterio(hoja, row, column, rango, evaluacion, tipoHoja)
                    Call superacionCriterio(hoja, row, column)
                Next
            Next
        End If
    Next
    hoja.Range("B5").Select
    Call protege(hoja)
End Sub
Sub estableceFormatoCondicional(hoja, row, column)
    Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
    hoja.Cells(row, column).FormatConditions.Delete
    Set cond3 = hoja.Cells(row, column).FormatConditions.Add(xlCellValue, xlEqual, "=-1")
    Set cond2 = hoja.Cells(row, column).FormatConditions.Add(xlCellValue, xlLess, "=0.5")
    Set cond1 = hoja.Cells(row, column).FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0.5")
    With cond1
        .Interior.Color = RGB(207, 237, 208)
        .Font.Color = RGB(43, 95, 23)
        .Font.Bold = True
    End With
    With cond2
        .Interior.Color = RGB(245, 201, 206)
        .Font.Color = RGB(140, 27, 21)
        .Font.Bold = True
    End With
    With cond3
        .Interior.Color = RGB(219, 225, 240)
        .Font.Color = RGB(219, 225, 240)
        .Font.Bold = False
    End With
End Sub
Sub calificacionCriterio(hoja, row, column, rango, evaluacion, tipoHoja)
    Dim bloqueado As Boolean
    Select Case tipoHoja
        Case miEvaluacion
            If rango = 1 Then
                With hoja.Cells(row, column - 1)
                    .Value = ""
                    .Locked = False
                End With
            Else
                With hoja.Cells(row, column - 1)
                    .FormulaLocal = calculaFormula(evaluacion, rango, row, column + 1)
                    .Locked = True
                End With
            End If
        Case miRecuperacion
            bloqueado = False
            If evaluacion = 2 And (rango = 2 Or (rango = 3 And row < 26)) Then
                ' En la segunda evaluaci—n protege lo que viene de la primera, en la 3» no hace falta por si recuperaci—n total
                ' La primera evaluaci—n va hasta el criterio 3.9 que est‡ en la fila 25
                bloqueado = True
            End If
            origen = evaluaciones(evaluacion)
            Formula = "=" + origen + "!" + col_letter(column - 1) + Trim(Str(row))
            With hoja.Cells(row, column - 1)
                .FormulaLocal = Formula
                .Locked = bloqueado
            End With
        Case miOtra
            If evaluacion = 1 Then
                origen = "Recu3"
            Else
                origen = "Ordinaria"
            End If
            Formula = "=" + origen + "!" + col_letter(column - 1) + Trim(Str(row))
            With hoja.Cells(row, column - 1)
                .FormulaLocal = Formula
                .Locked = False
            End With
    End Select
End Sub
Sub superacionCriterio(hoja, row, column)
    rowName = Trim(Str(row))
    ColumnName = col_letter(column - 1)
    ' Establece la f—rmula de c‡lculo del nivel de consecuci—n del criterio de evaluaci—n
    hoja.Cells(row, column).FormulaLocal = "=IF($B" + rowName + "," + ColumnName + rowName + "/$B" + rowName + ",-1)"
    '' A–ade formateo condicional a la celda
    Call estableceFormatoCondicional(hoja, row, column)
End Sub
Sub calificacionBloque(hoja, row, column, tope)
    ' Rellena la f—rmula del c‡lculo de la calificaci—n del bloque
    rangoPesos = "$B" + Trim(Str(row)) + ":$B" + tope
    rangoCalificaciones = col_letter(column - 1) + Trim(Str(row)) + ":" + col_letter(column - 1) + tope
    Formula = "=IF(SUM(" + rangoPesos + ")<>0,SUM(" + rangoCalificaciones + ")/SUM(" + rangoPesos + ")*$B" + Trim(Str(row - 1)) + "*10, " + Chr(34) + Chr(34) + ")"
    hoja.Cells(row - 1, column - 1).FormulaLocal = Formula
End Sub
Sub pesoCriterio(hoja, row, column, rango, evaluacion, tipoHoja)
    Select Case tipoHoja
        Case miEvaluacion
            If rango <> 1 Then
                With hoja.Cells(row, column - 2)
                    .FormulaLocal = calculaFormula(evaluacion, rango, row, column)
                    .NumberFormat = "General"
                End With
            Else
                ' Si estamos en el bloque 1 para las evaluaciones se borra  y desbloquea el peso del bloque
                With hoja.Cells(row, column - 2)
                    .Value = ""
                    .Locked = False
                End With
            End If
        Case miRecuperacion
            origen = evaluaciones(evaluacion)
            With hoja.Cells(row, column - 2)
                    .FormulaLocal = "=" + origen + "!" + col_letter(column - 2) + Trim(Str(row))
                    .NumberFormat = "General"
            End With
        Case miOtra
            If evaluacion = 1 Then
                origen = "Recu3"
            Else
                origen = "Ordinaria"
            End If
            With hoja.Cells(row, column - 2)
                .FormulaLocal = "=" + origen + "!" + col_letter(column - 2) + Trim(Str(row))
                .NumberFormat = "General"
            End With
    End Select
End Sub
Function calculaFormula(evaluacion, rango, fila, columna)
    ' Decide de donde toma los datos para cada evaluaci—n debido a la evaluaci—n continua
    Select Case evaluacion
        Case 1
            calculaFormula = formulaMax("1", columna - 2, fila - 3)
        Case 2
            If rango = 2 Or (rango = 3 And fila < 26) Then
                ' La primera evaluaci—n va hasta el criterio 3.9 que est‡ en la fila 25
                calculaFormula = "=Recu1!" + col_letter(columna - 2) + Trim(Str(fila))
            Else
                calculaFormula = formulaMax("2", columna - 2, fila - 12)
            End If
        Case 3
            If rango < 5 Then
                    calculaFormula = "=Recu2!" + col_letter(columna - 2) + Trim(Str(fila))
            Else
                    calculaFormula = formulaMax("3", columna - 2, fila - 55)
            End If
    End Select
End Function
Function formulaMax(eval As String, columna As Integer, fila As Integer) As String
    Dim form As String
    form = "=MAX("
    For numero = 1 To 3
        form = form + hojaExamenes + eval + Trim(Str(numero)) + "!" + col_letter(columna) + Trim(Str(fila)) + ","
    Next
    formulaMax = Left(form, Len(form) - 1) + ")"
End Function
Sub estableceRangosEvaluacion(ByRef rangos() As Integer)
    ' (evaluacion, #rango, inicio/fin)
    ' 1» Ev.
    rangos(1, 1, 1) = 5
    rangos(1, 1, 2) = 6
    rangos(1, 2, 1) = 8
    rangos(1, 2, 2) = 15
    rangos(1, 3, 1) = 17
    rangos(1, 3, 2) = 25
    rangos(1, 4, 1) = 0
    rangos(1, 4, 2) = 0
    rangos(1, 5, 1) = 0
    rangos(1, 5, 2) = 0
    rangos(1, 6, 1) = 0
    rangos(1, 6, 2) = 0
    ' 2» Ev.
    rangos(2, 1, 1) = 5
    rangos(2, 1, 2) = 6
    rangos(2, 2, 1) = 8
    rangos(2, 2, 2) = 15
    rangos(2, 3, 1) = 17
    rangos(2, 3, 2) = 37
    rangos(2, 4, 1) = 39
    rangos(2, 4, 2) = 58
    rangos(2, 5, 1) = 0
    rangos(2, 5, 2) = 0
    rangos(2, 6, 1) = 0
    rangos(2, 6, 2) = 0
    ' 3» Ev.
    rangos(3, 1, 1) = 5
    rangos(3, 1, 2) = 6
    rangos(3, 2, 1) = 8
    rangos(3, 2, 2) = 15
    rangos(3, 3, 1) = 17
    rangos(3, 3, 2) = 37
    rangos(3, 4, 1) = 39
    rangos(3, 4, 2) = 58
    rangos(3, 5, 1) = 60
    rangos(3, 5, 2) = 63
    rangos(3, 6, 1) = 65
    rangos(3, 6, 2) = 85
End Sub

