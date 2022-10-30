Attribute VB_Name = "Test"
Option Base 1
Const celdaEstado As String = "C6"
Public evaluaciones(3) As String
Public recuperaciones(3) As String
Public otras(2) As String
Const maxPreguntas = 10
Dim actual As Worksheet
Type pregunta
    linea As Integer
    peso As Double
End Type
' (evaluacion, examen # pregunta) decide en qué línea irá cada pregunta de cada examen y qué peso tendrá
Public preguntas(3, 3, maxPreguntas) As pregunta
Sub inicializaPreguntas(ByRef preguntas() As pregunta)
    For i = 1 To 3
        For j = 1 To 3
            For k = 1 To maxPreguntas
                preguntas(i, j, k).linea = 0
                preguntas(i, j, k).peso = 2#
            Next
        Next
    Next
    ' Examen 1 de la 1ª
    evaluacion = 1
    examen = 1
    preguntas(evaluacion, examen, 1).linea = 5
    preguntas(evaluacion, examen, 2).linea = 6
    preguntas(evaluacion, examen, 3).linea = 7
    preguntas(evaluacion, examen, 4).linea = 9
    preguntas(evaluacion, examen, 5).linea = 10
    ' Examen 2 de la 1ª
    evaluacion = 1
    examen = 2
    preguntas(evaluacion, examen, 1).linea = 14
    preguntas(evaluacion, examen, 1).peso = 1.5
    preguntas(evaluacion, examen, 2).linea = 15
    preguntas(evaluacion, examen, 2).peso = 1.5
    preguntas(evaluacion, examen, 3).linea = 16
    preguntas(evaluacion, examen, 3).peso = 1#
    preguntas(evaluacion, examen, 4).linea = 17
    preguntas(evaluacion, examen, 5).linea = 18
    preguntas(evaluacion, examen, 6).linea = 19
    preguntas(evaluacion, examen, 6).peso = 1#
    preguntas(evaluacion, examen, 7).linea = 20
    preguntas(evaluacion, examen, 7).peso = 1#
    ' Examen 1 de la 2ª
    evaluacion = 2
    examen = 1
    preguntas(evaluacion, examen, 1).linea = 14
    preguntas(evaluacion, examen, 2).linea = 15
    preguntas(evaluacion, examen, 3).linea = 16
    preguntas(evaluacion, examen, 4).linea = 17
    preguntas(evaluacion, examen, 5).linea = 19
    ' Examen 2 de la 2ª
    evaluacion = 2
    examen = 2
    preguntas(evaluacion, examen, 1).linea = 27
    preguntas(evaluacion, examen, 1).peso = 1.5
    preguntas(evaluacion, examen, 2).linea = 28
    preguntas(evaluacion, examen, 2).peso = 1.5
    preguntas(evaluacion, examen, 3).linea = 29
    preguntas(evaluacion, examen, 3).peso = 1#
    preguntas(evaluacion, examen, 4).linea = 31
    preguntas(evaluacion, examen, 5).linea = 32
    preguntas(evaluacion, examen, 6).linea = 33
    preguntas(evaluacion, examen, 6).peso = 1#
    preguntas(evaluacion, examen, 7).linea = 34
    preguntas(evaluacion, examen, 7).peso = 1#
    ' Examen 1 de la 3ª
    evaluacion = 3
    examen = 1
    preguntas(evaluacion, examen, 1).linea = 5
    preguntas(evaluacion, examen, 1).peso = 2.5
    preguntas(evaluacion, examen, 2).linea = 6
    preguntas(evaluacion, examen, 2).peso = 2.5
    preguntas(evaluacion, examen, 3).linea = 7
    preguntas(evaluacion, examen, 3).peso = 2.5
    preguntas(evaluacion, examen, 4).linea = 8
    preguntas(evaluacion, examen, 4).peso = 2.5
End Sub
Sub generaTest()
    Set actual = ActiveSheet
    actual.Range("b6").Select
    'Rellena las calificaciones de los nueve exámenes y pone marcas de agua
    Call rellenaCalificaciones
    'Pone marcas de agua a Evaluaciones
    Call estado("Procesa Evaluaciones")
    Call marcaAguaEvaluaciones
End Sub
Sub rellenaCalificaciones()
    Dim fileName As String
    Dim rangos(3, 3, 2) As Integer
    Dim libro As Workbook
    Dim hoja As Worksheet
    
    fileName = nombreTest(libroExamenes)
    Set libro = Workbooks.Open(fileName)
    Call inicializaRangosExamenes(rangos)
    Call inicializaPreguntas(preguntas)
    For evaluacion = 1 To 3
        For examen = 1 To 3
            nombreHoja = "Examen" & Trim(Str(evaluacion)) & Trim(Str(examen))
            Call estado("Procesando " & nombreHoja)
            Set hoja = libro.Worksheets(nombreHoja)
            Call desprotege(hoja)
            Call estableceCalificaciones(hoja, evaluacion, examen)
            Call ponMarcaEnHoja(hoja)
        Next
    Next
    libro.Save
    libro.Close
    Call estado("Fin de Proceso.")
End Sub
Sub estableceCalificaciones(hoja, evaluacion, examen)
    For pregunta_i = 1 To maxPreguntas
        linea = preguntas(evaluacion, examen, pregunta_i).linea
        peso = preguntas(evaluacion, examen, pregunta_i).peso
        If linea <> 0 Then
                ' Establece peso criterio
                Call establecePeso(hoja, linea, peso)
                Call estableceCalificacion(hoja, linea, peso)
        End If
    Next pregunta_i
End Sub
Sub establecePeso(hoja, linea, valor)
    hoja.Range("B" & Trim(Str(linea))).Value = valor
End Sub
Function calificaAleatoria(maximo)
    calificaAleatoria = (Int(Rnd * 11) / 10) * maximo
End Function
Sub estableceCalificacion(hoja, linea, peso)
    Dim alumno As Integer
    For alumno = 3 To 61 Step 2
        calif = calificaAleatoria(peso)
        rngAlumno = col_letter(alumno) + Trim(Str(linea))
        hoja.Range(rngAlumno).Value = calif
    Next
End Sub
Function nombreTest(nombre As String) As String
    nombreTest = ThisWorkbook.Path & Application.PathSeparator & directorioTest & Application.PathSeparator & nombre
End Function
Sub ponMarcaEnHoja(ByRef hoja As Worksheet)
    Call desprotege(hoja)
    Call marcaAgua(hoja)
    Call protege(hoja)
End Sub
Sub marcaAguaEvaluaciones()
    Dim fileName As String
    Dim rangos(3, 3, 2) As Integer
    Dim libro As Workbook
    Dim hoja As Worksheet
    
    Call inicializaHojasEvaluacion(evaluaciones, recuperaciones, otras)
    fileName = nombreTest(libroEvaluacion)
    Set libro = Workbooks.Open(fileName)
    For i = 1 To 3
            Set hoja = libro.Worksheets(evaluaciones(i))
            Call ponMarcaEnHoja(hoja)
            Set hoja = libro.Worksheets(recuperaciones(i))
            Call ponMarcaEnHoja(hoja)
            If i < 3 Then
                Set hoja = libro.Worksheets(otras(i))
                Call ponMarcaEnHoja(hoja)
            End If
    Next
    libro.Save
    libro.Close
    Call estado("Fin de Proceso.")
End Sub
Sub estado(ByVal texto As String)
    actual.Range(celdaEstado).Value = texto
End Sub
Public Sub marcaAgua(ByRef hoja As Worksheet)
    Dim StrIn As String
    StrIn = "T E S T"
    With hoja.Shapes.AddTextEffect(msoTextEffect9, StrIn, "Arial Black", 72#, msoFalse, msoFalse, 5, 5)
        .ScaleWidth 2.08, msoFalse, msoScaleFromTopLeft
        .ScaleHeight 1.23, msoFalse, msoScaleFromBottomRight
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.SchemeColor = 26
        .Fill.Transparency = 0.75
        .Shadow.Transparency = 0.5
        .Line.Visible = msoFalse
        'position at cell corner
        .Top = 1
        .Left = 290
    End With
End Sub
