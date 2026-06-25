Attribute VB_Name = "Module2"
Sub SendMailWAttachement()
    Dim path As String
    Dim email As String
    Dim cc As String
    Dim subject As String
    
    path = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D5").Value
    email = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D6").Value
    cc = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D7").Value
    subject = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D8").Value
    
    
    'MsgBox "C:\Users\user\Downloads\" + file_name
    
    'ActiveWorkbook.Worksheet("final").SaveAs ThisWorkbook.path & "\" & "final" & ".xlsm"    '<~~~  Change file extension as needed
    
    Dim OutApp As Object, OutMail As Object
    Dim NewWB As Workbook
    Dim TempFilePath As String, TempFileName As String, FileFullPath As String
    
    'Optimize performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Define the temporary path and filename
    TempFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" ' SAVES TO DESKTOP NU TE MAI COMPLICA CU FISIERE
    TempFileName = "Alerte logistica.xlsm"
    FileFullPath = TempFilePath & TempFileName
    
    ' Copy the sheet; automatically creates and activates a new Workbook.
    ThisWorkbook.Worksheets("final").Copy
    Set NewWB = ActiveWorkbook
    
    ' Save the new workbook and close it
    NewWB.SaveAs Filename:=FileFullPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    NewWB.Close SaveChanges:=False
    
    'Create and send the email
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next ' Prevents code crash if email fails
        With OutMail
            .To = email
            .subject = subject
            .cc = cc
            .Attachments.Add FileFullPath
            .Display
        End With
    On Error GoTo 0
    
    'Clean up the temporary file
    If Dir(FileFullPath) <> "" Then
        Kill FileFullPath
    End If
    
    ' Reset Excel settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Clear object variables
    Set OutMail = Nothing
    Set OutApp = Nothing
    Set NewWB = Nothing

End Sub



