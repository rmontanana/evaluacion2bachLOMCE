Attribute VB_Name = "Email"
Option Explicit

'Note: The macros in the other module use the functions below, do not change them.
'Do not forget to copy them into your own workbook

Function MacExcelWithMacMailPDFCatalinaAndUp(subject As String, mailbody As String, _
    toaddress As String, ccaddress As String, _
    bccaddress As String, displaymail As String, _
    attachment As String, otherattachments As String, _
    thesignature As String, thesender As String)
    ' Ron de Bruin 1-Feb-2021
    ' https://macexcel.com
    ' This script file is for macOS Catalina and higher and used in the VBA examples
    Dim ScriptStr As String, RunMyScript As String

    'Build the AppleScriptTask string
    ScriptStr = subject & ";" & mailbody & ";" & toaddress & ";" & ccaddress & ";" & _
        bccaddress & ";" & displaymail & ";" & attachment & ";" & otherattachments & ";" & thesignature & ";" & thesender
        
    'Call the RDBMacMailCatalinaAndUp.scpt script file with the AppleScriptTask function
    RunMyScript = AppleScriptTask("RDBMacMail.scpt", "CreateMailInCatalinaAndUp", CStr(ScriptStr))

    'Delete the pdf file we just mailed
   Kill attachment
End Function
Function CheckAppleScriptTaskExcelScriptFile(ScriptFileName As String) As Boolean
    'Function to Check if the AppleScriptTask script file exists
    'Ron de Bruin : 6-March-2016
    Dim AppleScriptTaskFolder As String
    Dim TestStr As String

    AppleScriptTaskFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    AppleScriptTaskFolder = Replace(AppleScriptTaskFolder, "/Desktop", "") & _
        "Library/Application Scripts/com.microsoft.Excel/"

    On Error Resume Next
    TestStr = Dir(AppleScriptTaskFolder & ScriptFileName, vbDirectory)
    On Error GoTo 0
    If TestStr = vbNullString Then
        CheckAppleScriptTaskExcelScriptFile = False
    Else
        CheckAppleScriptTaskExcelScriptFile = True
    End If
End Function
Function CreateFolderinMacOffice(NameFolder As String) As String
    'Function to create folder if it not exists in the Microsoft Office Folder
    'Ron de Bruin : 13-July-2020
    Dim OfficeFolder As String
    Dim PathToFolder As String
    Dim TestStr As String

    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") & _
        "Library/Group Containers/UBF8T346G9.Office/"

    PathToFolder = OfficeFolder & NameFolder

    On Error Resume Next
    TestStr = Dir(PathToFolder & "*", vbDirectory)
    On Error GoTo 0
    If TestStr = vbNullString Then
        MkDir PathToFolder
    End If
    CreateFolderinMacOffice = PathToFolder
End Function

Sub MailRangoPDF(rango, evaluacion, profesor, ByRef email_destino As String, ByRef email_origen As String)
    'Do not forget to also add the custom functions into your own workbook
    'More Information : https://macexcel.com/examples/mailpdf/macmail/
    'Ron de Bruin, 1-Feb-2021
    Dim FileName As String
    Dim FolderName As String
    Dim Folderstring As String
    Dim FilePathName As String
    Dim strbody As String
    Dim hoja As Worksheet

    'Check for AppleScriptTask script file that we must use to create the mail
    If CheckAppleScriptTaskExcelScriptFile(ScriptFileName:="RDBMacMail.scpt") = False Then
        MsgBox "Sorry the RDBMacMail.scpt is not in the correct location"
        Exit Sub
    End If

    Set hoja = Worksheets("inf_alumno")
    'If my ActiveSheet is landscape, I must attach this line
    'for making the PDF also landscape, seems to default to xlPortait
    hoja.PageSetup.Orientation = xlPortrait

    'Name of the folder in the Office folder
    FolderName = "inf_alumno"
    'Name of the pdf file, Date and Time in this example
    FileName = Format(Now, "dd-mmm-yyyy hh-mm-ss") & ".pdf"

    Folderstring = CreateFolderinMacOffice(NameFolder:=FolderName)
    FilePathName = Folderstring & Application.PathSeparator & FileName

    'Create the body text in the strbody string
    strbody = "Adjunto te env’o el informe correspondiente a la evaluaci—n: " & evaluacion & " de F’sica" _
        & vbNewLine & vbNewLine & "Saludos" & vbNewLine & vbNewLine & profesor
    'expression A variable that represents a Workbook, Sheet, Chart, or Range object.
    'the parameters are not working like in Excel for Windows
    hoja.Range(rango).ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
    FilePathName, Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, ignoreprintareas:=False

    'Call the MacExcelWithMacMailPDFCatalinaAndUp function to create the mail
    'When you use more mail addresses separate them with a ,
    'Change yes to no in the displaymail argument to send directly
    'You can attach other files also like "/Users/rondebruin/Desktop/YourFile.xlsx"
    'When you want to add more files separate each file path with with a ,
    'Look in Mail>Preferences for the name of the signature (you can use any signature in the signatures section)
    'Look in Mail>Preferences for the name of the mail account
    'Sender name (thesender) looks like this : "Your Name <your@mailaddress.com>"
    'Do not change the attachment argument in this PDF function
    
    MacExcelWithMacMailPDFCatalinaAndUp subject:="Informe de la evaluaci—n " & evaluacion, _
    mailbody:=strbody, _
    toaddress:=email_destino, _
    ccaddress:="", _
    bccaddress:="", _
    attachment:=FilePathName, _
    otherattachments:="", _
    displaymail:="no", _
    thesignature:="", _
    thesender:=email_origen
End Sub
