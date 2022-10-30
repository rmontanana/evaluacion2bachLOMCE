Attribute VB_Name = "Comun"
Public Const password As String = "patitofrito"
Public Const libroExamenes As String = "Examenes.xlsm"
Public Const libroEvaluacion As String = "Evaluacion.xlsm"
Public Const hojaExamenes = "[Examenes.xlsm]Examen"
Public Const directorioTest = "Test"
Public evaluaciones(3) As String
Public recuperaciones(3) As String
Public otras(2) As String
Function col_letter(column As Integer) As String
    col_letter = Split(Cells(1, column).Address, "$")(1)
End Function
Sub protege(ByRef hoja As Worksheet)
    Call hoja.Protect(password)
End Sub
Sub desprotege(ByRef hoja As Worksheet)
    Call hoja.Unprotect(password)
End Sub
Sub inicializa()
    ' Define separators and apply.
    Application.DecimalSeparator = "."
    Application.ThousandsSeparator = ","
    Application.UseSystemSeparators = False
    Application.Calculation = xlManual
End Sub
Sub finaliza()
    Application.Calculation = xlAutomatic
End Sub
Sub inicializaHojasEvaluacion(ByRef evaluaciones() As String, ByRef recuperaciones() As String, ByRef otras() As String)
    evaluaciones(1) = "Primera"
    evaluaciones(2) = "Segunda"
    evaluaciones(3) = "Tercera"
    recuperaciones(1) = "Recu1"
    recuperaciones(2) = "Recu2"
    recuperaciones(3) = "Recu3"
    otras(1) = "Ordinaria"
    otras(2) = "Extraordinaria"
End Sub
