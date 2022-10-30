Attribute VB_Name = "MacroModule"
Option Explicit
Sub SaveMailActiveSheetAsPDF()
    'Do not forget to also add the custom functions into your own workbook
    'More Information : https://macexcel.com/examples/mailpdf/macmail/
    'Ron de Bruin, 1-Feb-2021
    Dim FileName As String
    Dim FolderName As String
    Dim Folderstring As String
    Dim FilePathName As String
    Dim strbody As String

    'Check for AppleScriptTask script file that we must use to create the mail
    If CheckAppleScriptTaskExcelScriptFile(ScriptFileName:="RDBMacMail.scpt") = False Then
        MsgBox "Sorry the RDBMacMail.scpt is not in the correct location"
        Exit Sub
    End If

    'If my ActiveSheet is landscape, I must attach this line
    'for making the PDF also landscape, seems to default to xlPortait
    ActiveSheet.PageSetup.Orientation = ActiveSheet.PageSetup.Orientation

    'Name of the folder in the Office folder
    FolderName = "inf_alumno"
    'Name of the pdf file, Date and Time in this example
    FileName = Format(Now, "dd-mmm-yyyy hh-mm-ss") & ".pdf"

    Folderstring = CreateFolderinMacOffice(NameFolder:=FolderName)
    FilePathName = Folderstring & Application.PathSeparator & FileName

    'Create the body text in the strbody string
    strbody = "Hi there" & vbNewLine & vbNewLine & _
        "This is line 1" & vbNewLine & _
        "This is line 2" & vbNewLine & _
        "This is line 3" & vbNewLine & _
        "This is line 4"

    'expression A variable that represents a Workbook, Sheet, Chart, or Range object.
    'The parameters are not working like in Excel for Windows
    With ActiveSheet
        .PageSetup.Zoom = False
        .ExportAsFixedFormat _
        Type:=xlTypePDF, _
        FileName:=FilePathName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        ignoreprintareas:=False
    End With
    
    'Call the MacExcelWithMacMailPDFCatalinaAndUp function to create the mail
    'When you use more mail addresses separate them with a ,
    'Change yes to no in the displaymail argument to send directly
    'You can attach other files also like "/Users/rondebruin/Desktop/YourFile.xlsx"
    'When you want to add more files separate each file path with with a ,
    'Look in Mail>Preferences for the name of the signature (you can use any signature in the signatures section)
    'Look in Mail>Preferences for the name of the mail account
    'Sender name (thesender) looks like this : "Your Name <your@mailaddress.com>"
    'Do not change the attachment argument in this PDF function
    
    MacExcelWithMacMailPDFCatalinaAndUp subject:="This is a test macro to mail the ActiveSheet as PDF", _
    mailbody:=strbody, _
    toaddress:="ron@debruin.nl", _
    ccaddress:="", _
    bccaddress:="", _
    attachment:=FilePathName, _
    otherattachments:="", _
    displaymail:="yes", _
    thesignature:="", _
    thesender:=""
End Sub
Sub SaveMailActiveWorkbookAsPDF()
    'Do not forget to also add the custom functions into your own workbook
    'More Information : https://macexcel.com/examples/mailpdf/macmail/
    'Ron de Bruin, 1-Feb-2021
    Dim FileName As String
    Dim FolderName As String
    Dim Folderstring As String
    Dim FilePathName As String
    Dim strbody As String

    'Check for AppleScriptTask script file that we must use to create the mail
    If CheckAppleScriptTaskExcelScriptFile(ScriptFileName:="RDBMacMail.scpt") = False Then
        MsgBox "Sorry the RDBMacMail.scpt is not in the correct location"
        Exit Sub
    End If

    'If my ActiveSheet is landscape, I must attach this line
    'for making the PDF also landscape, seems to default to xlPortait
    'all sheets seems to follow the direction of the activesheet
    ActiveSheet.PageSetup.Orientation = ActiveSheet.PageSetup.Orientation

    'Name of the folder in the Office folder
    FolderName = "RDBMailTempFolder"
    'Name of the pdf file, Date and Time in this example
    FileName = Format(Now, "dd-mmm-yyyy hh-mm-ss") & ".pdf"

    Folderstring = CreateFolderinMacOffice(NameFolder:=FolderName)
    FilePathName = Folderstring & Application.PathSeparator & FileName

    'Create the body text in the strbody string
    strbody = "Hi there" & vbNewLine & vbNewLine & _
        "This is line 1" & vbNewLine & _
        "This is line 2" & vbNewLine & _
        "This is line 3" & vbNewLine & _
        "This is line 4"

    'expression A variable that represents a Workbook, Sheet, Chart, or Range object.
    'the parameters are not working like in Excel for Windows
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
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
    
    MacExcelWithMacMailPDFCatalinaAndUp subject:="This is a test macro to mail the whole workbook as PDF", _
    mailbody:=strbody, _
    toaddress:="ron@debruin.nl", _
    ccaddress:="", _
    bccaddress:="", _
    attachment:=FilePathName, _
    otherattachments:="", _
    displaymail:="yes", _
    thesignature:="", _
    thesender:=""
End Sub
Sub SaveMailRangeAsPDF()
    'Do not forget to also add the custom functions into your own workbook
    'More Information : https://macexcel.com/examples/mailpdf/macmail/
    'Ron de Bruin, 1-Feb-2021
    Dim FileName As String
    Dim FolderName As String
    Dim Folderstring As String
    Dim FilePathName As String
    Dim strbody As String

    'Check for AppleScriptTask script file that we must use to create the mail
    If CheckAppleScriptTaskExcelScriptFile(ScriptFileName:="RDBMacMail.scpt") = False Then
        MsgBox "Sorry the RDBMacMail.scpt is not in the correct location"
        Exit Sub
    End If

    'If my ActiveSheet is landscape, I must attach this line
    'for making the PDF also landscape, seems to default to xlPortait
    ActiveSheet.PageSetup.Orientation = ActiveSheet.PageSetup.Orientation

    'Name of the folder in the Office folder
    FolderName = "inf_alumno"
    'Name of the pdf file, Date and Time in this example
    FileName = Format(Now, "dd-mmm-yyyy hh-mm-ss") & ".pdf"

    Folderstring = CreateFolderinMacOffice(NameFolder:=FolderName)
    FilePathName = Folderstring & Application.PathSeparator & FileName

    'Create the body text in the strbody string
    strbody = "Hi there" & vbNewLine & vbNewLine & _
        "This is line 1" & vbNewLine & _
        "This is line 2" & vbNewLine & _
        "This is line 3" & vbNewLine & _
        "This is line 4"

    'expression A variable that represents a Workbook, Sheet, Chart, or Range object.
    'the parameters are not working like in Excel for Windows
    Range("A1:H6").ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
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
    
    MacExcelWithMacMailPDFCatalinaAndUp subject:="This is a test macro to mail the range A1:H13 as PDF", _
    mailbody:=strbody, _
    toaddress:="ron@debruin.nl", _
    ccaddress:="", _
    bccaddress:="", _
    attachment:=FilePathName, _
    otherattachments:="", _
    displaymail:="yes", _
    thesignature:="", _
    thesender:=""
End Sub

