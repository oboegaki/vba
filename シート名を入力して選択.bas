Attribute VB_Name = "Module1"
Option Explicit


Sub Run()
    On Error Resume Next
    Dim sheetName As String
    sheetName = InputBox("input sheetName")
    Sheets(sheetName).Select
End Sub





