Attribute VB_Name = "Examenes"
Sub Fill_Examenes()
    Dim libro As Workbook
    Dim actual As Worksheet
    
    Call inicializa
    Set actual = ActiveSheet
    actual.Range("B2").Select
    Set libro = Workbooks.Open(ThisWorkbook.Path & Application.PathSeparator & libroExamenes)
    Application.Calculation = xlManual
    contador = 0
    For evaluacion = 1 To 3
        For examen = 1 To 3
            nombre = "Examen" + Trim(Str(evaluacion)) + Trim(Str(examen))
            actual.Cells(2, 3).Value = "Procesando " + nombre
            DoEvents
            Call procesaHoja(actual, libro, nombre, evaluacion)
            contador = contador + 1
        Next
    Next
    actual.Cells(2, 3).Value = "Fin de Proceso."
    MsgBox "Se han procesado " + Str(contador) + " hojas"""
    Call finaliza
    libro.Protect password, True
    Call libro.Save
End Sub
Sub procesaHoja(actual, libro, nombre, evaluacion)
    Dim hoja As Worksheet
    Dim row As Integer, column As Integer
    Dim rangos(3, 3, 2) As Integer
    
    Call inicializaRangosExamenes(rangos)
    iniCol = 4
    finCol = 63
    Set hoja = libro.Worksheets(nombre)
    Call desprotege(hoja)
    ' Desbloquea la celda para la fecha del examen
    hoja.Range("B1:B2").Locked = False
    For rango = 1 To 3
        If rangos(evaluacion, rango, 1) <> 0 Then
            actual.Cells(2, 3).Value = "Procesando " + nombre + " Rango:" + Str(rango)
            DoEvents
            For column = iniCol To finCol Step 2
                For row = rangos(evaluacion, rango, 1) To rangos(evaluacion, rango, 2)
                    If column = iniCol Then
                        ' Desbloquea y borra la columna de pesos
                        With hoja.Cells(row, column - 2)
                            .Locked = False
                            .Value = ""
                        End With
                    End If
                    rowName = Trim(Str(row))
                    ColumnName = col_letter(column - 1)
                    ' Borra la casilla de la calificaci—n
                    With hoja.Cells(row, column - 1)
                        .Value = ""
                        .Locked = False
                    End With
                    ' Establece la f—rmula de c‡lculo del nivel de consecuci—n del criterio de evaluaci—n
                    hoja.Cells(row, column).FormulaLocal = "=IF($B" + rowName + "," + ColumnName + rowName + "/$B" + rowName + ",-1)"
                    ' A–ade formateo condicional a la celda
                    Call estableceFormato(hoja, row, column)
                Next
            Next
        End If
    Next
    hoja.Range("B5").Select
    Call protege(hoja)
End Sub
Sub inicializaRangosExamenes(ByRef rangos() As Integer)
    '(evaluacion, #rango, inicio/fin)
    ' 1» Ev.
    rangos(1, 1, 1) = 5
    rangos(1, 1, 2) = 12
    rangos(1, 2, 1) = 14
    rangos(1, 2, 2) = 34
    rangos(1, 3, 1) = 0
    rangos(1, 3, 2) = 0
    ' 2» Ev.
    rangos(2, 1, 1) = 5
    rangos(2, 1, 2) = 25
    rangos(2, 2, 1) = 27
    rangos(2, 2, 2) = 46
    rangos(2, 3, 1) = 0
    rangos(2, 3, 2) = 0
    ' 3» Ev.
    rangos(3, 1, 1) = 5
    rangos(3, 1, 2) = 8
    rangos(3, 2, 1) = 10
    rangos(3, 2, 2) = 30
    rangos(3, 3, 1) = 0
    rangos(3, 3, 2) = 0
End Sub
Sub estableceFormato(hoja, row, column)
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

