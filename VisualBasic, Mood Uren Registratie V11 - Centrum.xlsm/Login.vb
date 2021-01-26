VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "LoginForm"
   ClientHeight    =   2610
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4350
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoginButton1_Click()
    Dim L1 As Variant
    Dim L2 As Variant
    Dim L3 As Variant
    Dim X As Integer
    
    Dim myDate As Date
    myDate = Now()

    L1 = TextBox1.value                                    'Username
    L2 = TextBox2.value                                    'Password

    Sheets("INDEX").Select

    If L1 = "timderoos" Then                               '\\ User = tim

        If L2 = "timderoos" Then                           '\\ Login procedure for Tim
            Range("B15").value = "Mood Eindhoven, Tim de Roos"
            Sheets("ADMIN").Range("B7").value = "Tim de Roos"
            Sheets("ADMIN").Range("B9").value = myDate
            For X = 1 To 4
                'Debug.Print X
                ActiveSheet.Shapes.Range(Array("Menu" & X)).Visible = True
            Next
            Unload Me
        Else
            Label5.ForeColor = &HFF&                       '\\ Wrong password
        End If

    ElseIf L1 = "willemaarts" Then                         '\\ User = Willem
    
        If L2 = "willemaarts" Then                         '\\ Login procedure for Willem
            Range("B15").value = "Mood Eindhoven, Willem Aarts"
            Sheets("ADMIN").Range("B7").value = "Willem Aarts"
            Sheets("ADMIN").Range("B9").value = myDate
            For X = 1 To 4
                'Debug.Print X
                ActiveSheet.Shapes.Range(Array("Menu" & X)).Visible = True
            Next
            Application.Run "INDEXSpecs.ADMIN"
            Unload Me
        Else
            Label5.ForeColor = &HFF&                       '\\ Wrong password
        End If

    ElseIf L1 = "Streetfood" Then                          '\\ User = Streetfood
    
        If L2 = "Streetfood" Then                          '\\ Login procedure for User
            Range("B15").value = "Mood Streetfood, User"
            Sheets("ADMIN").Range("B7").value = "User"
            Sheets("ADMIN").Range("B9").value = myDate
            For X = 1 To 4
                'Debug.Print X
                ActiveSheet.Shapes.Range(Array("Menu" & X)).Visible = True
            Next
            Unload Me
        Else
            Label5.ForeColor = &HFF&                       '\\ Wrong password
        End If
    
    Else
        Label5.ForeColor = &HFF&                           '\\ Wrong username
    End If

    Range("A4").Select
    
    Debug.Print "Username; " & L1
    Debug.Print "Password; " & L2
End Sub

Private Sub CancelButton1_Click()
    Unload Me
End Sub

Private Sub LoginButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    With LoginButton1
        .BackStyle = fmBackStyleOpaque
        .BackColor = &HC0C0C0
    End With
    
End Sub

Private Sub CancelButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CancelButton1.BackStyle = fmBackStyleOpaque
    CancelButton1.BackColor = &HC0C0C0
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LoginButton1.BackStyle = fmBackStyleTransparent
    CancelButton1.BackStyle = fmBackStyleTransparent
    'LoginButton1.BackColor = RGB(255, 255, 255)
End Sub

Private Sub UserForm_Initialize()

    #If Mac Then
        ResizeUserForm Me
    #End If
    
    Label5.ForeColor = &HFFFFFF

End Sub

