Attribute VB_Name = "Global_Var"
'Used for Public variables with employee form
Option Explicit

Public Fdate As String
Public Ldate As String

Sub Application_On()
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub

Sub Application_Off()
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
End Sub

