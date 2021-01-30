VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteEmployee 
   Caption         =   "Delete Employee"
   ClientHeight    =   2805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4650
   OleObjectBlob   =   "DeleteEmployee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub EmplDelete1_Click()
    Dim answer As String
    Dim Foundcell As Range
    Dim E1 As Variant, E2 As Variant
    
    If EmpName1.value = "" Then Exit Sub
    
    answer = MsgBox("Do you want to delete employee;" & vbNewLine & vbNewLine & EmpName1.value & vbNewLine & vbNewLine & " From the Excel file?", _
                    vbCritical + vbYesNoCancel, "Delete employee?")
                    
    'Debug.Print answer
    
    Select Case answer
    
    Case vbYes
        Debug.Print "vbYes"
    Case vbNo
        Debug.Print "vbNo"
        Exit Sub
    Case vbCancel
        Debug.Print "vbCancel"
        Exit Sub
    End Select
    
    Debug.Print "Start process -> Delete employee; " & EmpName1
    
    Application.Run ("Global_Var.Application_Off")
    
    '\\ Find employee in DataStr sheet
    Sheets("DataStr").Select
    
    Set Foundcell = Range("A:A").Find(What:=EmpName1)
    If Not Foundcell Is Nothing Then
        'MsgBox (EmpName1 & " Found in row: " & FoundCell.Row)
    Else
        MsgBox (EmpName1 & "; Was not found in Sheets(DataStr). Deleting process had ended.")
        Sheets("INDEX").Select
        Application.Run ("Global_Var.Application_On")
        Exit Sub
    End If
        
    E1 = Foundcell.Row
    Debug.Print EmpName1.value & " was found in row; " & E1
    
    With CheckBox1
        .value = True
        .Caption = "Sheet; DataStr | " & EmpName1.value
    End With
    
    '\\ Find employee in DataStr sheet
    Sheets("DataEmp").Select
    
    Set Foundcell = Range("A:A").Find(What:=EmpName1)
    If Not Foundcell Is Nothing Then
        'MsgBox (EmpName1 & " Found in row: " & FoundCell.Row)
    Else
        MsgBox (EmpName1 & "; Was not found in Sheets(DataEmp). Deleting process had ended.")
        Sheets("INDEX").Select
        Application.Run ("Global_Var.Application_On")
        Exit Sub
    End If
        
    E2 = Foundcell.Row
    Debug.Print EmpName1.value & " was found in row; " & E2
    
    With CheckBox2
        .value = True
        .Caption = "Sheet; DataEmp | " & EmpName1.value
    End With
    
    Sheets("INDEX").Select
    
    If CheckBox1.value And CheckBox2.value = True Then
        Debug.Print CheckBox1.value & " " & CheckBox2.value & "; ON!"
    Else
        MsgBox "In 1 of the 2 sheets the employee name has not been found, so therefore the process was terminated."
        Application.Run ("Global_Var.Application_On")
        Exit Sub
    End If
    
    '\\ Start deleting proces HERE
    '\\ Look for idea for archiving data from the deleted employee
    
    
    Application.Run ("Global_Var.Application_On")
    
End Sub

Private Sub EmpName1_Change()
    EmplDelete1.Enabled = True
End Sub

Private Sub UserForm_Initialize()
    Dim c As Range
    Dim n As Long

    #If Mac Then
        ResizeUserForm Me
    #End If

    With Me
        .CheckBox1.value = False
        .CheckBox2.value = False
    End With

    Sheets("DataEmp").Select                               'Kijkt hoeveel medewerkers er zijn
    n = Cells(1, 1).End(xlDown).Row

    For Each c In Sheets("DataEmp").Range("A2:A" & n)      'Zet alle namen in de combobox
        Me.EmpName1.AddItem c.value
    Next

    Sheets("INDEX").Select
    
    EmplDelete1.Enabled = False
    
End Sub


