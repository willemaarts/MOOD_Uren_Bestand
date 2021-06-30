VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmployeeForm 
   Caption         =   "Employee Card"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6630
   OleObjectBlob   =   "EmployeeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EmployeeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dateperiod1_Click()
    'Voor YELLOW, GREEN, RED, BLUE = Zie OneNote voor aanteking (aantal spaties in datum notatie)

    Dim L1 As Variant                                      'begin datum
    Dim L2 As Variant                                      'Eind datum
    Dim N1 As Variant                                      'Naam medewerker
    Dim Foundcell As Range
    Dim dateVariable As Variant

    Dim F1 As Variant, F2 As Variant, F3 As Variant        'wordt gebruikt voor selected cells
    
    If EmpName1.value = "" Then
        MsgBox "Geen naam gekozen, probeer opnieuw."
        Exit Sub
    End If

    On Error Resume Next                                   'Openend Calendar en kiest 2 datums en zet die op de index sheet
    dateVariable = CalendarForm.GetDate( _
                   OkayButton:=True, _
                   ShowWeekNumbers:=True)
    
    Debug.Print Fdate
    Debug.Print Ldate
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    N1 = EmpName1.value                                    'Onthoudt naam Medewerker
    
    Sheets("INDEX").Select
    
    L1 = Fdate                                             'Public variables from sepperated module
    L2 = Ldate
    
    Sheets("DataStr").Select
    
    Dim VarianceDate As String: VarianceDate = L1          'Begin datum zoeken

    Dim TargetCell As Range, TargetCol As Integer
    Set TargetCell = Rows("1").Find(What:=CDate(VarianceDate), LookIn:=xlFormulas, LookAt:=xlPart)
    If Not TargetCell Is Nothing Then
        TargetCol = TargetCell.Column
        ' MsgBox TargetCol
    Else
        MsgBox (L1 & " Not found. Please try again, and select 2 valid dates")
        Sheets("INDEX").Select
        Exit Sub
    End If

    F1 = TargetCell.Column                                 'Locatie voor Eerste datum
        
    VarianceDate = L2

    Set TargetCell = Rows("1").Find(What:=CDate(VarianceDate), LookIn:=xlFormulas, LookAt:=xlPart)
    If Not TargetCell Is Nothing Then
        TargetCol = TargetCell.Column
        ' MsgBox TargetCol
    Else
        MsgBox (L2 & " Not found. Please try again, and select 2 valid dates")
        Sheets("INDEX").Select
        Exit Sub
    End If

    F2 = TargetCell.Column                                 'Locatie voor Laatste datum
    
    Set Foundcell = Range("A:A").Find(What:=N1)            'Locatie medewerker
    If Not Foundcell Is Nothing Then
        'MsgBox (N1 & " Found in row: " & FoundCell.Row)
    Else
        MsgBox (N1 & " not found")
        Sheets("INDEX").Select
        Exit Sub
    End If
        
    F3 = Foundcell.Row                                     'onthoud de locatie van de medewerker
        
    Range(Cells(F3, F1), Cells(F3, F2 + 3)).Select         'Selecteerd de cellen (vd Tijden) die worden verplaatst
                           
    Dim SelRange As Variant                                'onthoudt de locatie van de geselecteerde cellen (vd Tijden)
    SelRange = Selection.Address(ReferenceStyle:=xlA1, _
                                 RowAbsolute:=False, ColumnAbsolute:=False)
    
    SelRange = "=DataStr!" & SelRange                      'Maakt de formule gereed voor EMP blad
    
    Sheets("<EMP>").Select
    Range("A2").Select
    
    ActiveCell.Formula = SelRange                          'zet de formule in de cell naar de verwijzing van de orginele tijden
    
    On Error Resume Next
    Range("A2").Select                                     'Komt opeens een '@' in de formule staan en dit zorgt ervoor dat die weg gaat
    ActiveCell.Replace What:="@", Replacement:="", LookAt:=xlPart, _
                       SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                       ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    [2:2].value = [2:2].value                              'Maakt een hard copy van de cellen
    
    Sheets("DataStr").Select                               'hier ga je terug om de datums te kopieren
    Range(Cells(1, F1), Cells(1, F2 + 3)).Select           'hier selecteer je de datums van de gekozen tijden
    
    'Onthoudt de locatie van de geselecteerde cellen (vd Datums)
    SelRange = Selection.Address(ReferenceStyle:=xlA1, _
                                 RowAbsolute:=False, ColumnAbsolute:=False)

    SelRange = "=DataStr!" & SelRange                      'Maakt de formule gereed voor EMP blad
    
    Sheets("<EMP>").Select
    Range("A1").Select
   
    ActiveCell.Formula = SelRange                          'zet de formule in de cell naar de verwijzing van de orginele tijden
    
    On Error Resume Next
    Range("A1").Select                                     'Komt opeens een '@' in de formule staan en dit zorgt ervoor dat die weg gaat
    ActiveCell.Replace What:="@", Replacement:="", LookAt:=xlPart, _
                       SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                       ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
   
    [1:1].value = [1:1].value                              'Maakt een hard copy van de cellen
    Rows("1:1").Select                                     'Hier zorg je ervoor dat de datum de juiste format heeft
    Selection.NumberFormat = "m/d/yyyy"
   
    Range("A5").value = N1                                 'noteert de naam van de medewerker
   
    Unload Me
   
    Sheets("TEMP-MTseven").Select                          'kopieerd de standaard sheet naar een Temp blad
    ActiveSheet.Copy After:=Sheets(Sheets.count)
    ActiveSheet.Select

    ActiveSheet.Range("F3").value = N1                     'noteert de naam van de medewerker
    ActiveSheet.Name = ActiveSheet.Range("F3")             'geeft de sheet de naam van de medewerker

    [B2:K61].value = [B2:K61].value                        'Maakt een hard copy van de cellen
    
    Sheets("<EMP>").Select                                 'verwijderd de gegevens voor volgend gebruik
    Rows("1:5").Select
    Selection.ClearContents
    
    Sheets(N1).Select                                      'gaat terug naar de sheet
    Range("F3").Select
    Application.Run ("HourOverview.NAWemp")                'gaat de NAW gegevens ophalen
    
    
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

End Sub

Private Sub EmplSearch1_Click()
    Dim L1 As Variant, L2 As Variant, L3 As Variant
    Dim Foundcell As Range

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    Sheets("DataEmp").Select

    L1 = EmpName1.value                                    'Onthoud de naam van de medewerker

    Set Foundcell = Range("A:A").Find(What:=L1)            'Zoekt de naam in het werknemersblad
    If Not Foundcell Is Nothing Then
        'MsgBox (L1 & " Found in row: " & FoundCell.Row)
    Else
        MsgBox (L1 & " not found")
        Sheets("INDEX").Select
        Exit Sub
    End If

    TextBox3.value = Cells(Foundcell.Row, 1).value         'Noteert de naam van de medewerker
    TextBox2.value = Cells(Foundcell.Row, 2).value         'Noteert het telefoonnummer
    TextBox1.value = Cells(Foundcell.Row, 3).value         'noteert het email adres
    TextBox5.value = Cells(Foundcell.Row, 9).value         'Noteert de geboortedatum
    TextBox6.value = Cells(Foundcell.Row, 10).value        'Noteert de loon
    ComboBox4.value = Cells(Foundcell.Row, 11).value       'Noteert de vestiging

    If Cells(Foundcell.Row, 4).value = "Ja" Then           'Wanneer de medewerker een shirt heeft
        CheckBox1.value = True                             'Wordt de checkbox aangevinkt
    Else
        CheckBox1.value = False
    End If

    If Cells(Foundcell.Row, 6).value = "Gelezen" Then      'Zelfde als bij shirt maar nu
        CheckBox2.value = True                             'bij voorschrift
    Else
        CheckBox2.value = False
    End If
    
    ComboBox2.value = Cells(Foundcell.Row, 5).value        'Noteert het aantal shirts
    
    ComboBox3.value = Cells(Foundcell.Row, 7).value        'Noteert het soort contract
    
    TextBox4.value = Cells(Foundcell.Row, 8).value         'Noteert de eventuele opmerking
    
    If CheckBox1.value = False Then                        'Wanneer er geen shirt is uitgeleend, gaat het aantal
        ComboBox2.value = "0"                              'shirt automatisch op "0"
    End If

    Sheets("INDEX").Select
    
    Dateperiod1.Enabled = True
    Change1.Enabled = True

    Range("C7").value = L1                                 'Zet de naam van de medewerker op het hoofdblad zodat wanneer
    With Application                                       'je de datum wilt opzoeken dat dan de naam bekend is.
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

End Sub

Private Sub Cancel1_Click()
    Unload Me
End Sub

Private Sub Change1_Click()
    Dim L1 As Variant, L2 As Variant, L3 As Variant
    Dim Foundcell As Range

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    If EmpName1.value = "" Then
        MsgBox "No name selected"
        Exit Sub
    End If

    Sheets("DataEmp").Select

    L1 = EmpName1.value                                    'noteert de naam van de medewerker

    Set Foundcell = Range("A:A").Find(What:=L1)            'Zoekt de medewerker op in het medewerkersblad
    If Not Foundcell Is Nothing Then
        'MsgBox (L1 & " Found in row: " & FoundCell.Row)
    Else
        MsgBox (L1 & " not found")
        Sheets("INDEX").Select
        Exit Sub
    End If

    Cells(Foundcell.Row, 1).value = TextBox3.value         'Veranderd de naam van de medewerker
    Cells(Foundcell.Row, 2).value = TextBox2.value         'Veranderd het telefoon nummer
    Cells(Foundcell.Row, 3).value = TextBox1.value         'veranderd het email adres
    Cells(Foundcell.Row, 9).value = TextBox5.value         'Veranderd het geboorte datum
    Cells(Foundcell.Row, 10).value = TextBox6.value        'veranderd het loon
    Cells(Foundcell.Row, 11).value = ComboBox4.value       'veranderd de vestiging

    If CheckBox1.value = True Then                         'Wanneer de checkbox is aangevinkt veranderd de cel
        Cells(Foundcell.Row, 4).value = "Ja"               'naar ja!
    Else
        Cells(Foundcell.Row, 4).value = "Nee"
    End If

    If CheckBox2.value = True Then                         'zelfde als shirt, maar dan met voorschrift
        Cells(Foundcell.Row, 6).value = "Gelezen"
    Else
        Cells(Foundcell.Row, 6).value = "Nee"
    End If

    Cells(Foundcell.Row, 5).value = ComboBox2.value        'veranderd aantal shirts
    
    Cells(Foundcell.Row, 7).value = ComboBox3.value        'veranderd soort contract
    
    Cells(Foundcell.Row, 8).value = TextBox4.value         'veranderd eventuele opmerking
    
    Range("I1").Select

    Sheets("INDEX").Select

    Range("C7").value = L1                                 'noteert de naam op het voorblad

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

    MsgBox "Info changed"

    Call EmplSearch1_Click

End Sub

Private Sub UserForm_Initialize()
    'Add emptyrow till search
    Dim c As Range
    Dim n As Long
    Dim X As Variant

    #If Mac Then
        ResizeUserForm Me
    #End If

    Sheets("DataEmp").Select                               'Kijkt hoeveel medewerkers er zijn
    n = Cells(1, 1).End(xlDown).Row

    For Each c In Sheets("DataEmp").Range("A2:A" & n)      'Zet alle namen in de combobox
        Me.EmpName1.AddItem c.value
    Next

    Sheets("INDEX").Select

    For X = 1 To 10                                        'dit is voor het aantal shirts
        Me.ComboBox2.AddItem X
    Next

    Me.ComboBox3.AddItem "0-uren contract"
    Me.ComboBox3.AddItem "38-uren contract"

    Me.ComboBox4.AddItem "Mood Eindhoven"
    Me.ComboBox4.AddItem "Mood Strijp-s"
    Me.ComboBox4.AddItem "Mood Streetfood"

    Me.ComboBox4.AddItem "Mood Eindhoven - Keuken"
    Me.ComboBox4.AddItem "Mood Strijp-S - Keuken"
    Me.ComboBox4.AddItem "Mood Streetfood - Keuken"
    
    Change1.Enabled = False
    Dateperiod1.Enabled = False
    
End Sub


