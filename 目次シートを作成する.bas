Attribute VB_Name = "Module1"
Option Explicit


Sub Run()
    Dim sh As Worksheet
    Dim index As String
    Dim count As Integer
    
    index = "INDEX"
    count = 0
    
    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = index
    
    For Each sh In Sheets
        count = count + 1
        Sheets(index).Range("A" & count).Value = sh.Name
    Next

End Sub





