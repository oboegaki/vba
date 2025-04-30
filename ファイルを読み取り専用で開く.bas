Attribute VB_Name = "Module1"
Option Explicit


Sub Run()

    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim path As String
    
    path = "hoge"
    Set wb = Workbooks.Open(path, ReadOnly:=True)
    
    'close
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
    
ErrorHandler:
    ' close
    If Not wb Is Nothing Then
        Application.DisplayAlerts = False
        wb.Close
        Application.DisplayAlerts = True
    End If
    
    MsgBox "error"
End Sub





