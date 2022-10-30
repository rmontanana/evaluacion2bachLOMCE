Attribute VB_Name = "Informes"
Option Explicit
Const botonEditar = "button 8"
Const botonOcultar = "button 9"
Const ReportsFolder = "Informes"
Const maxRow = 85
Dim campos(3) As String
Type configuracion
    fontsize As Integer
    asize As Integer
    csize As Integer
    dsize As Integer
End Type
Sub Generar()
    Dim tipo As String
    Dim alumno_idx As Integer
    Dim i As Integer
    
    DesprotegeLibro
    Application.DisplayAlerts = False
    alumno_idx = Range("Alumno").Value - 1
    tipo = Range("Tipos").Item(Range("Tipo"), 1)
    If alumno_idx = 0 Then
        For i = 1 To 30
            informeAlumno i, tipo
        Next
    Else
        informeAlumno alumno_idx, tipo
    End If
    Application.DisplayAlerts = True
    ProtegeLibro
End Sub
Sub informeAlumno(idAlumno, tipoInforme)
    Dim alumno As String
    Dim row As Integer
    Dim evaluacion As String
    Dim columna_origen As String
    Dim rango As Range
    Dim rangotxt As String
    Dim rangoOrigen As String
    Dim informe As Worksheet
    Dim origen As Worksheet
    Dim profesor As String
    Dim email_destino As String
    Dim email_origen As String
    Dim conf As configuracion
    Dim FilePathName As String

    ' Coge los datos de entrada
    alumno = Range("Alumnos").Item(idAlumno + 1, 1)
    evaluacion = Range("Evaluaciones").Item(Range("Evaluacion"), 1)
    profesor = Range("Profesor")
    email_origen = Range("email")
    email_destino = Range("Emails").Item(idAlumno)
    ' Inicializa hoja del informe
    Set informe = Worksheets("inf_alumno")
    informe.Visible = xlSheetVisible
    Set origen = Worksheets(evaluacion)
    With informe.Range("inf_all")
        .ClearContents
        .ClearFormats
    End With
    ' Copia Nombre del alumno y la evaluaci뾫 a la hoja del informe
    informe.Range("inf_nombre") = alumno
    With informe.Range("c1:d2")
        .Merge across:=False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(219, 225, 240)
        .Font.Color = RGB(45, 78, 116)
        .Font.Size = 20
        .Font.Bold = True
        .WrapText = True
    End With
    informe.Range("inf_evaluacion") = evaluacion
    ' Copia texto de criterios de evaluaci뾫 de la hoja de evaluaci뾫 elegida al informe
    rangotxt = calculaRango("A1", evaluacion, "B")
    origen.Range(rangotxt).Copy
    informe.Range(rangotxt).PasteSpecial Paste:=xlPasteAll
    Range("A1") = "F뭩ica " & Range("A1")
    ' Copia calificaciones de la hoja de evaluacion뾫 elegida al informe
    '   Calcula el rango del alumno a imprimir
    '  Calificaciones
    columna_origen = col_letter((idAlumno * 2) + 1)
    rangoOrigen = calculaRango(columna_origen & "3", evaluacion, col_letter((idAlumno * 2) + 2))
    rangotxt = calculaRango("C3", evaluacion, "D")
    origen.Range(rangoOrigen).Copy
    informe.Range(rangotxt).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    origen.Range(rangoOrigen).Copy
    informe.Range(rangotxt).PasteSpecial Paste:=xlPasteFormats
    ' Formato de calificaci뾫
    With informe.Range("C3:D3")
        .Merge across:=False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 20
    End With
    ' Restaura el formato condicional al porcentaje de consecuci뾫 del criterio
    For row = 4 To maxRow
        estableceFormatoCondicional informe, row, 4
    Next
    ' Modifica el formato de salida
    conf = estableceConfiguracion(evaluacion)
    With informe
        .Range("C3").Font.Size = conf.fontsize
        .Columns("A").ColumnWidth = conf.asize
        .Columns("C").ColumnWidth = conf.csize
        .Columns("D").ColumnWidth = conf.dsize
    End With
    rangotxt = calculaRango("A4", evaluacion, "D")
    With informe.Range(rangotxt)
        .Font.Size = conf.fontsize
    End With
    ' General el informe
    rangotxt = calculaRango("A1", evaluacion, "D")
    Select Case tipoInforme
        Case "Impresora":
            informe.PrintOut Preview:=True, from:=1, To:=2, ignoreprintareas:=True
        Case "PDF":
            FilePathName = ActiveWorkbook.path & Application.PathSeparator & ReportsFolder
            If NoExisteDir(FilePathName) Then
                MkDir FilePathName
            End If
            FilePathName = FilePathName & Application.PathSeparator & nombreArchivo(evaluacion, alumno)
            informe.Range(rangotxt).ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
                FilePathName, Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, ignoreprintareas:=False
        Case Else:
            MailRangoPDF rangotxt, evaluacion, profesor, email_destino, email_origen
    End Select
    informe.Visible = xlSheetHidden
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
Function col_letter(column As Integer) As String
    col_letter = Split(Cells(1, column).Address, "$")(1)
End Function
Function NoExisteDir(path)
On Error Resume Next
    ChDir path
    If Err Then NoExisteDir = True Else NoExisteDir = False
End Function
Function nombreArchivo(evaluacion, alumno)
    Dim nom As String
    Dim alumno_filtrado As String
    alumno_filtrado = Replace(alumno, " ", "_")
    alumno_filtrado = Replace(alumno_filtrado, ",", "")
    Select Case evaluacion
        Case "Primera"
            nom = "1틿v"
        Case "Recu1"
            nom = "Recup_1틿v"
        Case "Segunda"
            nom = "2틿v"
        Case "Recu2"
            nom = "Recup_2틿v"
        Case "Tercera"
            nom = "3틿v"
        Case "Recu3"
            nom = "Recup_3틿v"
        Case Else
            nom = evaluacion
    End Select
    nombreArchivo = nom & "_" & alumno_filtrado & ".pdf"
End Function
Function estableceConfiguracion(evaluacion) As configuracion
    Dim res As configuracion
    Select Case evaluacion
        Case "Primera", "Recu1":
            res.asize = 80
            res.fontsize = 16
            res.csize = 10
            res.dsize = 7
        Case "Segunda", "Recu2":
            res.asize = 80
            res.fontsize = 14
            res.csize = 10
            res.dsize = 7
        Case Else:
            res.asize = 160
            res.fontsize = 22
            res.csize = 10
            res.dsize = 10
    End Select
    estableceConfiguracion = res
End Function
Function calculaRango(origen, evaluacion, columna2)
    Dim resultado As String
    
    resultado = origen & ":"
    Select Case evaluacion
        Case "Primera", "Recu1":
            resultado = resultado & columna2 & "25"
        Case "Segunda", "Recu2":
            resultado = resultado & columna2 & "58"
        Case Else:
            resultado = resultado & columna2 & Trim(Str(maxRow))
    End Select
    calculaRango = resultado
End Function
Sub inicia()
    campos(1) = "Servidor"
    campos(2) = "usuario"
    campos(3) = "passwd"
End Sub
Sub EditConfig()
    Dim rango As Range
    Dim i As Integer
    ActiveSheet.Shapes(botonOcultar).Visible = True
    ActiveSheet.Shapes(botonEditar).Visible = False
    Desprotege
    inicia
    For i = 1 To 3
        With Range(campos(i))
            .Interior.Color = RGB(207, 237, 208)
            .Font.Color = RGB(43, 95, 23)
            .Font.Bold = True
        End With
        Set rango = Range(campos(i))
        rango.MergeArea.Locked = False
    Next
    Protege
End Sub
Sub HideConfig()
    Dim rango As Range
    Dim i As Integer
    inicia
    Desprotege
    ActiveSheet.Shapes(botonOcultar).Visible = False
    ActiveSheet.Shapes(botonEditar).Visible = True
    For i = 1 To 3
        With Range(campos(i))
            .Interior.Color = RGB(217, 217, 217)
            .Font.Color = RGB(217, 217, 217)
            .Font.Bold = True
        End With
        Set rango = Range(campos(i))
        rango.MergeArea.Locked = True
    Next
    Protege
End Sub
