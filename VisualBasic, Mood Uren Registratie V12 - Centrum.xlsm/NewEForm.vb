VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewEForm 
   Caption         =   "Add new employee"
   ClientHeight    =   6510
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4020
   OleObjectBlob   =   "NewEForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewEForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox1_Change()

    If ComboBox1.value = "Ja" Then                         'wanneer er shirts in bruikleen zijn, dan komt
        Label12.Left = 150
    End If
    
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Dim emptyRow As Long

    If TextBox1.value = "" Then
        MsgBox "Geen naam ingevoerd"
        Exit Sub
    End If

    If ComboBox5.value = "" Then
        MsgBox "Geen vestiging gekozen"
        Exit Sub
    End If
    
    Application.Run ("Global_Var.Application_Off")

    Sheets("DataEmp").Select
    emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1  'pakt de onderste vrije regel

    Cells(emptyRow, 1).value = TextBox1.value              'Noteert de naam
    Cells(emptyRow, 2).value = TextBox3.value              'noteert het telefoonnummer
    Cells(emptyRow, 3).value = TextBox2.value              'noteert het emailadres
    Cells(emptyRow, 4).value = ComboBox1.value             'noteert of ze een shirt hebben
    Cells(emptyRow, 5).value = ComboBox2.value             'noteert het aantal shirts
    Cells(emptyRow, 6).value = ComboBox3.value             'noteert voorschrift
    Cells(emptyRow, 7).value = ComboBox4.value             'noteert soort contract
    Cells(emptyRow, 9).value = TextBox6.value              'noteert geboortedatum
    Cells(emptyRow, 10).value = TextBox7.value             'noteert loon
    Cells(emptyRow, 11).value = ComboBox5.value            'noteert vestiging

    Cells(emptyRow, 9).NumberFormat = "m/d/yyyy"
    Cells(emptyRow, 10).Style = "Currency"

    Sheets("DataStr").Select
    emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1  'pakt de onderste vrije regel
    Cells(emptyRow, 1).value = TextBox1.value              'Noteert de naam
    Cells(emptyRow, 2).value = ComboBox5.value             'noteert vestiging

    Sheets("INDEX").Select
    
    Application.Run ("Global_Var.Application_On")

    MsgBox "Nieuwe medewerker genoteerd; " & TextBox1.value
    
    Unload Me
    NewEForm.Show

End Sub

Private Sub UserForm_Initialize()
    Dim X As Variant

    #If Mac Then
        ResizeUserForm Me
    #End If

    Me.ComboBox1.AddItem "Ja"
    Me.ComboBox1.AddItem "Nee"

    For X = 1 To 10
        Me.ComboBox2.AddItem X
    Next

    Me.ComboBox3.AddItem "Gelezen"
    Me.ComboBox3.AddItem "Nee"

    Me.ComboBox4.AddItem "0-uren contract"
    Me.ComboBox4.AddItem "38-uren contract"

    Me.ComboBox5.AddItem "Mood Eindhoven"
    Me.ComboBox5.AddItem "Mood Strijp-S"
    Me.ComboBox5.AddItem "Mood Streetfood"

    Me.ComboBox5.AddItem "Mood Eindhoven - Keuken"
    Me.ComboBox5.AddItem "Mood Strijp-S - Keuken"
    Me.ComboBox5.AddItem "Mood Streetfood - Keuken"

End Sub

