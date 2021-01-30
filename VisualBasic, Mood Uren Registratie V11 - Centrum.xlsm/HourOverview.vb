Attribute VB_Name = "HourOverview"
Option Explicit
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/
'/Made by ItWill - Willem Aarts - willemaarts@itwill.nl
'/
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub NAWemp()
    Dim L1 As Variant
    Dim L2 As Variant
    Dim L3 As Variant
    Dim L4 As Variant

    Dim C1 As Variant
    Dim C2 As Variant

    Dim Foundcell As Range

    L1 = Range("F3").value                                 'Onthoud de naam van de medewerker

    Sheets("DataEmp").Select
    
    Set Foundcell = Range("A:A").Find(What:=L1)            'Zoekt de naam in het werknemersblad
    If Not Foundcell Is Nothing Then
        'MsgBox (L1 & " Found in row: " & FoundCell.Row)
    Else
        MsgBox (L1 & " not found")
        Exit Sub
    End If

    L2 = Cells(Foundcell.Row, 1).value                     'Noteert de naam van de medewerker
    L3 = Cells(Foundcell.Row, 2).value                     'Noteert het telefoonnummer
    L4 = Cells(Foundcell.Row, 3).value                     'noteert het email adres

    C1 = Cells(Foundcell.Row, 11).value                    'noteert de vestiging

    If C1 = "Mood Eindhoven" Then
        C2 = "Eindhoven@mood.nl"
        
    ElseIf C1 = "Mood Eindhoven - Keuken" Then
        C2 = "Eindhoven@mood.nl"
        
    ElseIf C1 = "Mood Strijp-s" Then
        C2 = "strijps@mood.nl"

    ElseIf C1 = "Mood Strijp-s - Keuken" Then
        C2 = "strijps@mood.nl"

    ElseIf C1 = "Mood Streetfood" Then
        C2 = "Streetfood@mood.nl"
        
    ElseIf C1 = "Mood Streetfood - Keuken" Then
        C2 = "Streetfood@mood.nl"
        
    Else
        MsgBox "Vestiging verkeerd"
    End If
    
    Sheets(L1).Select
    Range("F4").value = L3
    Range("F5").value = L4
    
    Range("F7").value = C1
    Range("F8").value = C2

End Sub

