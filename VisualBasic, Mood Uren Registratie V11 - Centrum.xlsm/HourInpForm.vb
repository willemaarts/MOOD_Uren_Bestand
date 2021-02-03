VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HourInpForm 
   Caption         =   "Uren registratie"
   ClientHeight    =   6690
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5520
   OleObjectBlob   =   "HourInpForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HourInpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EntBtn_Click()
    Dim Dte As Date                                        'onthoud de datum
    Dim Foundcell As Range
    Dim cell As Range
    Dim c As Variant
    Dim Q1 As Variant, Q2 As Variant

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    c = Emp1.value
    
    If Emp1.value = "" Then
        MsgBox "Geen naam geselecteerd."
        Exit Sub
    End If

    Sheets("DataStr").Select
    
    Dte = Date1.value                                      'Gewerkte uren

    Dim VarianceDate As String: VarianceDate = Dte         'Begin datum zoeken

    Dim TargetCell As Range, TargetCol As Integer
    Set TargetCell = Rows("1").Find(What:=CDate(VarianceDate), LookIn:=xlFormulas, LookAt:=xlPart)
    If Not TargetCell Is Nothing Then
        TargetCol = TargetCell.Column
        ' MsgBox TargetCol
    Else
        MsgBox (Dte & " not found")
        Sheets("INDEX").Select
        Exit Sub
    End If

    Q1 = TargetCell.Column                                 'Locatie voor Eerste datum
    
    Set Foundcell = Range("A:A").Find(What:=c)             'Zoekt de naam op
    If Not Foundcell Is Nothing Then
        'MsgBox (c & " Found in row: " & FoundCell.Row)
    Else
        MsgBox (c & " not found")
        Exit Sub
    End If
    
    Q2 = Foundcell.Row                                     'locatie naam
    
    Range(Cells(Q2, Q1), Cells(Q2, Q1)).Select             'Selecteert de juiste datum

    If CheckBox1.value = True Then
        ActiveCell.value = A1
        ActiveCell.Offset(0, 1).value = A2
        ActiveCell.Offset(0, 2).value = A3
        ActiveCell.Offset(0, 3).value = A4
    End If

    If CheckBox2.value = True Then
        ActiveCell.Offset(0, 4).value = B1
        ActiveCell.Offset(0, 5).value = B2
        ActiveCell.Offset(0, 6).value = B3
        ActiveCell.Offset(0, 7).value = B4
    End If
    
    If CheckBox3.value = True Then
        ActiveCell.Offset(0, 8).value = C1
        ActiveCell.Offset(0, 9).value = C2
        ActiveCell.Offset(0, 10).value = C3
        ActiveCell.Offset(0, 11).value = C4
    End If
    
    If CheckBox4.value = True Then
        ActiveCell.Offset(0, 12).value = D1
        ActiveCell.Offset(0, 13).value = D2
        ActiveCell.Offset(0, 14).value = D3
        ActiveCell.Offset(0, 15).value = D4
    End If
    
    If CheckBox5.value = True Then
        ActiveCell.Offset(0, 16).value = E1
        ActiveCell.Offset(0, 17).value = E2
        ActiveCell.Offset(0, 18).value = E3
        ActiveCell.Offset(0, 19).value = E4
    End If
    
    If CheckBox6.value = True Then
        ActiveCell.Offset(0, 20).value = F1
        ActiveCell.Offset(0, 21).value = F2
        ActiveCell.Offset(0, 22).value = F3
        ActiveCell.Offset(0, 23).value = F4
    End If
    
    If CheckBox7.value = True Then
        ActiveCell.Offset(0, 24).value = G1
        ActiveCell.Offset(0, 25).value = G2
        ActiveCell.Offset(0, 26).value = G3
        ActiveCell.Offset(0, 27).value = G4
    End If
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

    Sheets("INDEX").Select
    MsgBox ("Uren genoteerd bij; " & c)

    Unload Me                                              'restart het uren notatie
    HourInpForm.Show
    
End Sub

Private Sub CalcBtn_Click()
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    Dim Total As Date, Total_1 As Date, Total_2 As Date, Total_3 As Date, Total_4 As Date
    Dim T_1 As Variant, T_2 As Variant, T_3 As Variant, T_4 As Variant, T_5 As Variant, T_6 As Variant, T_7 As Variant
    Dim T_8 As Variant

    If CheckBox1.value = True Then
        If A2.value < "06:00" Then
            Total_1 = ((TimeValue(A1) - TimeValue("23:59"))) '+ TimeValue("00:01"))
            Total_2 = ((TimeValue(A2) - TimeValue("00:01"))) ' + TimeValue("00:01"))
            Total_3 = (TimeValue(Total_1) + TimeValue(Total_2))
            Total_4 = (TimeValue(Total_3) - TimeValue(A3))
            If A2.value = "00:00" Then
                Total = (TimeValue(Total_4))
            Else
                Total = (TimeValue(Total_4) + TimeValue("00:02"))
            End If
            A4.value = Format(Total, "hh:mm")
            
        Else
            Total_1 = TimeValue(A1) - TimeValue(A2)
            Total = TimeValue(Total_1) - TimeValue(A3)
        
            A4.value = Format(Total, "hh:mm")
            
        End If
        T_1 = Total
    Else
        T_1 = TimeValue("00:00")
    End If
     
    If CheckBox2.value = True Then
        If B2.value < "06:00" Then
            Total_1 = ((TimeValue(B1) - TimeValue("23:59"))) '+ TimeValue("00:01"))
            Total_2 = ((TimeValue(B2) - TimeValue("00:01"))) ' + TimeValue("00:01"))
            Total_3 = (TimeValue(Total_1) + TimeValue(Total_2))
            Total_4 = (TimeValue(Total_3) - TimeValue(B3))
            If B2.value = "00:00" Then
                Total = (TimeValue(Total_4))
            Else
                Total = (TimeValue(Total_4) + TimeValue("00:02"))
            End If
            B4.value = Format(Total, "hh:mm")
            
        Else
            Total_1 = TimeValue(B1) - TimeValue(B2)
            Total = TimeValue(Total_1) - TimeValue(B3)
        
            B4.value = Format(Total, "hh:mm")
            
        End If
        T_2 = Total
    Else
        T_2 = TimeValue("00:00")
    End If


    If CheckBox3.value = True Then
        If C2.value < "06:00" Then
            Total_1 = ((TimeValue(C1) - TimeValue("23:59"))) '+ TimeValue("00:01"))
            Total_2 = ((TimeValue(C2) - TimeValue("00:01"))) ' + TimeValue("00:01"))
            Total_3 = (TimeValue(Total_1) + TimeValue(Total_2))
            Total_4 = (TimeValue(Total_3) - TimeValue(C3))
            If C2.value = "00:00" Then
                Total = (TimeValue(Total_4))
            Else
                Total = (TimeValue(Total_4) + TimeValue("00:02"))
            End If
            C4.value = Format(Total, "hh:mm")
            
        Else
            Total_1 = TimeValue(C1) - TimeValue(C2)
            Total = TimeValue(Total_1) - TimeValue(C3)
        
            C4.value = Format(Total, "hh:mm")
            
        End If
        T_3 = Total
    Else
        T_3 = TimeValue("00:00")
    End If


    If CheckBox4.value = True Then
        If D2.value < "06:00" Then
            Total_1 = ((TimeValue(D1) - TimeValue("23:59"))) '+ TimeValue("00:01"))
            Total_2 = ((TimeValue(D2) - TimeValue("00:01"))) ' + TimeValue("00:01"))
            Total_3 = (TimeValue(Total_1) + TimeValue(Total_2))
            Total_4 = (TimeValue(Total_3) - TimeValue(D3))
            If D2.value = "00:00" Then
                Total = (TimeValue(Total_4))
            Else
                Total = (TimeValue(Total_4) + TimeValue("00:02"))
            End If
            D4.value = Format(Total, "hh:mm")
            
        Else
            Total_1 = TimeValue(D1) - TimeValue(D2)
            Total = TimeValue(Total_1) - TimeValue(D3)
        
            D4.value = Format(Total, "hh:mm")
            
        End If
        T_4 = Total
    Else
        T_4 = TimeValue("00:00")
    End If


    If CheckBox5.value = True Then
        If E2.value < "06:00" Then
            Total_1 = ((TimeValue(E1) - TimeValue("23:59"))) '+ TimeValue("00:01"))
            Total_2 = ((TimeValue(E2) - TimeValue("00:01"))) ' + TimeValue("00:01"))
            Total_3 = (TimeValue(Total_1) + TimeValue(Total_2))
            Total_4 = (TimeValue(Total_3) - TimeValue(E3))
            If E2.value = "00:00" Then
                Total = (TimeValue(Total_4))
            Else
                Total = (TimeValue(Total_4) + TimeValue("00:02"))
            End If
            E4.value = Format(Total, "hh:mm")
            
        Else
            Total_1 = TimeValue(E1) - TimeValue(E2)
            Total = TimeValue(Total_1) - TimeValue(E3)
        
            E4.value = Format(Total, "hh:mm")
            
        End If
        T_5 = Total
    Else
        T_5 = TimeValue("00:00")
    End If


    If CheckBox6.value = True Then
        If F2.value < "06:00" Then
            Total_1 = ((TimeValue(F1) - TimeValue("23:59"))) '+ TimeValue("00:01"))
            Total_2 = ((TimeValue(F2) - TimeValue("00:01"))) ' + TimeValue("00:01"))
            Total_3 = (TimeValue(Total_1) + TimeValue(Total_2))
            Total_4 = (TimeValue(Total_3) - TimeValue(F3))
            If F2.value = "00:00" Then
                Total = (TimeValue(Total_4))
            Else
                Total = (TimeValue(Total_4) + TimeValue("00:02"))
            End If
            F4.value = Format(Total, "hh:mm")
            
        Else
            Total_1 = TimeValue(F1) - TimeValue(F2)
            Total = TimeValue(Total_1) - TimeValue(F3)
        
            F4.value = Format(Total, "hh:mm")
            
        End If
        T_6 = Total
    Else
        T_6 = TimeValue("00:00")
    End If


    If CheckBox7.value = True Then
        If G2.value < "06:00" Then
            Total_1 = ((TimeValue(G1) - TimeValue("23:59"))) '+ TimeValue("00:01"))
            Total_2 = ((TimeValue(G2) - TimeValue("00:01"))) ' + TimeValue("00:01"))
            Total_3 = (TimeValue(Total_1) + TimeValue(Total_2))
            Total_4 = (TimeValue(Total_3) - TimeValue(G3))
            If G2.value = "00:00" Then
                Total = (TimeValue(Total_4))
            Else
                Total = (TimeValue(Total_4) + TimeValue("00:02"))
            End If
            G4.value = Format(Total, "hh:mm")

        Else
            Total_1 = TimeValue(G1) - TimeValue(G2)
            Total = TimeValue(Total_1) - TimeValue(G3)
        
            G4.value = Format(Total, "hh:mm")

        End If
        T_7 = Total
    Else
        T_7 = TimeValue("00:00")
    End If

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

    Range("A4").Select

    EntBtn.Enabled = True

End Sub

Private Sub CancBtn_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    #If Mac Then
        ResizeUserForm Me
    #End If
    
    Application.Run ("Global_Var.Application_Off")

    A4 = Format(A4.value, "hh:mm")
    B4 = Format(A4.value, "hh:mm")
    C4 = Format(A4.value, "hh:mm")
    D4 = Format(A4.value, "hh:mm")
    E4 = Format(A4.value, "hh:mm")
    F4 = Format(A4.value, "hh:mm")
    G4 = Format(A4.value, "hh:mm")

    Dim c As Range
    Dim n As Long
    
    Sheets("DataStr").Select                               'kijkt hoeveel medewerkers er zijn
    n = Cells(1, 2).End(xlDown).Row
    
    For Each c In Sheets("DataStr").Range("A2:A" & n)      'zet de namen in de ComboBox
        Me.Emp1.AddItem c.value
    Next
    
    Sheets("INDEX").Select
    
    EntBtn.Enabled = False                                 'zorgt ervoor dat de knop niet zomaar kan worden ingedrukt
    
    Dim i As Integer
    Dim myDate As Date
    myDate = Now()                                         'Zet de datum klaar in de ComboBox
    For i = -18 To 0                                       'Add the next 15 days, for example
        Date1.AddItem Format(DateAdd("d", i, myDate), "dd/mm/yyyy")
    Next
    Date1.ListIndex = 11
    
    Dim T As Double                                        'Dit zorgt ervoor dat de tijden worden ingevuld
    
    For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
        A1.AddItem Format(T, "hh:mm")
    Next T
    'ComboBox2.ListIndex = 72

    For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
        A2.AddItem Format(T, "hh:mm")
    Next T
    'ComboBox3.ListIndex = 94
    
    For T = TimeValue("12:00 AM") To TimeValue("01:14 AM") Step TimeSerial(0, 15, 0)
        A3.AddItem Format(T, "hh:mm")
    Next T
    'ComboBox4.ListIndex = 0
    
    
    Dim ctrl As Control

    For Each ctrl In Me.Controls
        If TypeName(ctrl) <> "ComboBox" Then
        Else
            If ctrl.Name = "Emp1" Then
            Else
                If ctrl.Name = "Date1" Then
                Else
                    'MsgBox Ctrl.Name
                
                    If ctrl.Name = "B1" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                    Else
                    End If
                

                    If ctrl.Name = "B2" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox3.ListIndex = 94
                    Else
                    End If
                
                
                    If ctrl.Name = "B3" Then
                        For T = TimeValue("12:00 AM") To TimeValue("01:14 AM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox4.ListIndex = 0
                    Else
                    End If
                
                
                    If ctrl.Name = "C1" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                    Else
                    End If
                

                    If ctrl.Name = "C2" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox3.ListIndex = 94
                    Else
                    End If
                
                
                    If ctrl.Name = "C3" Then
                        For T = TimeValue("12:00 AM") To TimeValue("01:14 AM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox4.ListIndex = 0
                    Else
                    End If
            
            
                    If ctrl.Name = "D1" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                    Else
                    End If
                

                    If ctrl.Name = "D2" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox3.ListIndex = 94
                    Else
                    End If
                
                
                    If ctrl.Name = "D3" Then
                        For T = TimeValue("12:00 AM") To TimeValue("01:14 AM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox4.ListIndex = 0
                    Else
                    End If
                
                
                    If ctrl.Name = "E1" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                    Else
                    End If
                

                    If ctrl.Name = "E2" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox3.ListIndex = 94
                    Else
                    End If
                
                
                    If ctrl.Name = "E3" Then
                        For T = TimeValue("12:00 AM") To TimeValue("01:14 AM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox4.ListIndex = 0
                    Else
                    End If
                
                
                    If ctrl.Name = "F1" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                    Else
                    End If
                

                    If ctrl.Name = "F2" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox3.ListIndex = 94
                    Else
                    End If
                
                
                    If ctrl.Name = "F3" Then
                        For T = TimeValue("12:00 AM") To TimeValue("01:14 AM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox4.ListIndex = 0
                    Else
                    End If
                
                
                    If ctrl.Name = "G1" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                    Else
                    End If
                

                    If ctrl.Name = "G2" Then
                        For T = TimeValue("12:00 AM") To TimeValue("11:45 PM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox3.ListIndex = 94
                    Else
                    End If
                
                
                    If ctrl.Name = "G3" Then
                        For T = TimeValue("12:00 AM") To TimeValue("01:14 AM") Step TimeSerial(0, 15, 0)
                            ctrl.AddItem Format(T, "hh:mm")
                        Next T
                        'ComboBox4.ListIndex = 0
                    Else
                    End If
                
                End If
            End If
        End If
    Next ctrl
    
    Application.Run ("Global_Var.Application_On")
    
End Sub

Private Sub CheckBox1_Click()
    If CheckBox1.value = True Then
        A1.ListIndex = 46                                  '72
        A2.ListIndex = 80                                  '94
        A3.ListIndex = 4                                   '0
    Else
        A1.value = ""
        A2.value = ""
        A3.value = ""
        A4.value = ""
    End If
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2.value = True Then
        B1.ListIndex = 46
        B2.ListIndex = 80
        B3.ListIndex = 4
    Else
        B1.value = ""
        B2.value = ""
        B3.value = ""
        B4.value = ""
    End If
End Sub

Private Sub CheckBox3_Click()
    If CheckBox3.value = True Then
        C1.ListIndex = 46
        C2.ListIndex = 80
        C3.ListIndex = 4
    Else
        C1.value = ""
        C2.value = ""
        C3.value = ""
        C4.value = ""
    End If
End Sub

Private Sub CheckBox4_Click()
    If CheckBox4.value = True Then
        D1.ListIndex = 46
        D2.ListIndex = 80
        D3.ListIndex = 4
    Else
        D1.value = ""
        D2.value = ""
        D3.value = ""
        D4.value = ""
    End If
End Sub

Private Sub CheckBox5_Click()
    If CheckBox5.value = True Then
        E1.ListIndex = 46
        E2.ListIndex = 80
        E3.ListIndex = 4
    Else
        E1.value = ""
        E2.value = ""
        E3.value = ""
        E4.value = ""
    End If
End Sub

Private Sub CheckBox6_Click()
    If CheckBox6.value = True Then
        F1.ListIndex = 46
        F2.ListIndex = 80
        F3.ListIndex = 4
    Else
        F1.value = ""
        F2.value = ""
        F3.value = ""
        F4.value = ""
    End If
End Sub

Private Sub CheckBox7_Click()
    If CheckBox7.value = True Then
        G1.ListIndex = 46
        G2.ListIndex = 80
        G3.ListIndex = 4
    Else
        G1.value = ""
        G2.value = ""
        G3.value = ""
        G4.value = ""
    End If
End Sub

