Attribute VB_Name = "Module1"
Option Explicit


Sub Run()

    Dim eMail As String
    eMail = "abcdeffff@excel233333.com"

    '@Å@len
    MsgBox InStr(eMail, "@")


    '"@" before len
    MsgBox Left(eMail, InStr(eMail, "@") - 1)


    '"@" after len
    MsgBox Right(eMail, Len(eMail) - InStr(eMail, "@") - 1)

End Sub





