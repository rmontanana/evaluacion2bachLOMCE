Attribute VB_Name = "Comun"
Sub muestraAlumnos(alumnos As Range)
    alumnos.Columns.Select
    Selection.EntireColumn.Hidden = False
    Range("A1").Select
End Sub
Sub Protege()
    ActiveSheet.Protect ("patitofrito"), DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFiltering:=True
End Sub
Sub Desprotege()
    ActiveSheet.Unprotect ("patitofrito")
End Sub
Public Sub procesaDobleClick(ByVal target As Range)
    Dim rngAlumnos As Range
    Dim rInt As Range
    ' Oculta/Muestra las columnas de alumnos para obtener informes.
    Set rngAlumnos = Range("C1:BJ30") ' Rango de nœmeros de alumno
    Set rInt = Intersect(target, rngAlumnos)
    If Not rInt Is Nothing Then
        Desprotege
        rngAlumnos.Columns.Select
        If Selection.Columns(1).Hidden Or Selection.Columns(2).Hidden Or Selection.Columns(3).Hidden Then
            Call muestraAlumnos(rngAlumnos)
        Else
            Selection.EntireColumn.Hidden = True
            target.Columns.Select
            Selection.EntireColumn.Hidden = False
        End If
        Protege
    End If
End Sub

