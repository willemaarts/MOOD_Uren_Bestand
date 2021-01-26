VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WeaklyForm 
   Caption         =   "Wekelijkse urenstaat"
   ClientHeight    =   5805
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4650
   OleObjectBlob   =   "WeaklyForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WeaklyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdEnt_Click()
    Dim Q1 As Variant, Q2 As Variant, Q3 As Variant        'Vestiging
    Dim Q4 As Variant, Q5 As Variant, Q6 As Variant        'Vestiging

    Dim P1 As Variant, P2 As Variant                       'fooi berekenen

    Dim Tt1 As Variant, Tt2 As Variant                     'Totalen uren en fooi

    Dim cell As Range                                      'Om de locatie van de datum te zoeken
    Dim Dte As Date                                        'Onthoudt de Datum als Date variabel

    Dim X As Integer                                       'Gebruikt voor de For Loop
    Dim Column As Long                                     'Gebruikt voor de values goed te maken

    Dim emptyRow As Long
    Dim L1 As Variant                                      'onthoudt de Column

    Dte = DD1 & "-" & MM1 & "-" & YYYY1

    If Fooi1.value = "" Then
        Fooi1.value = "00.000001"
    End If

    If M1.value = True Then
        Q1 = "Mood Eindhoven"
        Q2 = "Mood Streetfood"
        Q3 = "Mood Strijp-s"
    End If

    If K1.value = True Then
        Q4 = "Mood Eindhoven - Keuken"
        Q5 = "Mood Streetfood - Keuken"
        Q6 = "Mood Strijp-s - Keuken"
    End If

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    Sheets("DataStr").Select

    Dim VarianceDate As String: VarianceDate = Dte         '\\Begin datum zoeken

    Dim TargetCell As Range, TargetCol As Integer
    Set TargetCell = Rows("1").Find(What:=CDate(VarianceDate), LookIn:=xlFormulas, LookAt:=xlPart)
    If Not TargetCell Is Nothing Then
        TargetCol = TargetCell.Column
        ' MsgBox TargetCol
    Else
        MsgBox (Dte & " not found")
        With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
        End With
        Exit Sub
    End If

    L1 = TargetCell.Column - 1                             '\\Locatie voor Eerste datum

    Sheets("DataStr").Copy After:=Sheets(Sheets.count)
    With ActiveSheet
        .Select
        .Name = "WEAKLYTEMP"
    End With

    Columns("B:B").Select
    Selection.AutoFilter
    ActiveSheet.Range("B:B").AutoFilter Field:=1, Criteria1:=Array(Q1, Q2, Q3, Q4, Q5, Q6), Operator:=xlFilterValues
    Range("B1").Select

    Dim oRow As Range, rng As Range                        '\\ Vanaf hier worden die vestigingen die niet gekozen zijn verwijderd
    Dim myRows As Range
    
    With Sheets("WEAKLYTEMP")
        Set myRows = Intersect(.Range("A:A").EntireRow, .UsedRange)
        If myRows Is Nothing Then Exit Sub
    End With

    For Each oRow In myRows.Columns(1).Cells
        If oRow.EntireRow.Hidden Then
            If rng Is Nothing Then
                Set rng = oRow
            Else
                Set rng = Union(rng, oRow)                 '\\Weet niet of dit goed werkt op macbook
            End If
        End If
    Next

    If Not rng Is Nothing Then rng.EntireRow.Delete

    Range("A1").Select

    Range(Columns(2), Columns(L1)).Delete Shift:=xlToLeft  '\\ Alle datums die verwijderen die er niet toe doen

    Columns("AD:AD").Select                                '\\verwijderd alle data na de gekozen periode
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    
    Application.CutCopyMode = False

    On Error GoTo Error_Value
    With Range("AE2")                                      '\\Hier speciale error handling
        .Select
        .FormulaR1C1 = _
                     "=(RC[-26]+RC[-22]+RC[-18]+RC[-14]+RC[-10]+RC[-6]+RC[-2])*24"
        .NumberFormat = "General"
    End With
    
    emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1  '\\Hier kan Error ontstaan, wanneer er geen Checkboxes zijn aangeklikt
    Selection.AutoFill Destination:=Range("AE2:AE" & emptyRow), Type:=xlFillDefault '\\Idem

    Range("AE" & emptyRow + 1).Select

    '|||||||||||||||||||||||||||||||||| ERROR FUNCTIE KLAAR||||||||||||||||||||||||
    On Error GoTo 0
    
    On Error Resume Next
    With Selection
        .Formula = "=IFERROR(SUM(AE2:AE" & emptyRow - 1 & "),1)" '\\Change last '1' to other value for better errorhandler for user
    End With
    Debug.Print Err.Number
    
    If Err.Number > 0 Then '1004
        Selection.Formula = "=IFERROR(SUM(AE2:AE" & emptyRow - 1 & ");1)" '\\Change last '1' to other value for better errorhandler for user
    End If

    Selection.NumberFormat = "0.00"
    
    On Error GoTo 0
    Debug.Print Err.Number
    '|||||||||||||||||||||||||||||||||| ERROR FUNCTIE KLAAR||||||||||||||||||||||||


    If ActiveCell.value = "0" Then                         '\\Total Hour Worked
        P1 = "00.0000001"
    Else
        P1 = ActiveCell.value
    End If
    
    Tt1 = ActiveCell.value                                 '\\ Totaal uren

    ActiveCell.Offset(0, 1).Formula = "=sum(AF2:AF" & emptyRow + 1 & ")" '\\ SUM functie Fooi

    P2 = Fooi1 / P1                                        '\\Fooi per Medewerker per Uur
    With Range("AG2")
        .value = P2
        .NumberFormat = "0.00"
    End With

    Tt2 = ActiveCell.Offset(0, 1).value                                '\\ Totaal fooi

    With Range("AF2")
        .FormulaR1C1 = "=RC[-1]*R2C33"                     '\\Berekening van Fooi (AG2 met vaste waarden)
        .NumberFormat = "0.00"
        .Select
    End With

    Selection.AutoFill Destination:=Range("AF2:AF" & emptyRow), Type:=xlFillDefault

    Application.CutCopyMode = False

    Sheets("TEMP-WEAKLY").Select
    
    Range("B1").Formula2R1C1 = "Name; "
    Range("C1").Formula2R1C1 = "=WEAKLYTEMP!R1C2"          '\\ Datum
    Range("D1").Formula2R1C1 = "=WEAKLYTEMP!R1C6"
    Range("E1").Formula2R1C1 = "=WEAKLYTEMP!R1C10"
    Range("F1").Formula2R1C1 = "=WEAKLYTEMP!R1C14"
    Range("G1").Formula2R1C1 = "=WEAKLYTEMP!R1C18"
    Range("H1").Formula2R1C1 = "=WEAKLYTEMP!R1C22"
    Range("I1").Formula2R1C1 = "=WEAKLYTEMP!R1C26"
    
    Range("K1").Formula2R1C1 = "Uur Totaal"
    Range("M1").Formula2R1C1 = "Fooi"
    
    Range("B2").FormulaR1C1 = "=WEAKLYTEMP!RC1"            '\\Naam

    Range("C2").FormulaR1C1 = "=WEAKLYTEMP!RC5"            '\\Tijd Ma
    Range("D2").FormulaR1C1 = "=WEAKLYTEMP!RC9"            '\\Tijd Di
    Range("E2").FormulaR1C1 = "=WEAKLYTEMP!RC13"           '\\Tijd Wo
    Range("F2").FormulaR1C1 = "=WEAKLYTEMP!RC17"           '\\Tijd Do
    Range("G2").FormulaR1C1 = "=WEAKLYTEMP!RC21"           '\\Tijd Vri
    Range("H2").FormulaR1C1 = "=WEAKLYTEMP!RC25"           '\\Tijd Za
    Range("I2").FormulaR1C1 = "=WEAKLYTEMP!RC29"           '\\Tijd Zo

    Range("K2").FormulaR1C1 = "=WEAKLYTEMP!RC31"           '\\Uur totaal
    Range("M2").FormulaR1C1 = "=WEAKLYTEMP!RC32"           '\\Fooi per Persoon

    Range("B2:M2").Select
    Selection.AutoFill Destination:=Range("B2:M" & emptyRow - 1), Type:=xlFillDefault
    
    Range("B" & emptyRow + 1).Select

    With ActiveCell
        .Select
        .FormulaR1C1 = "Totaal = "
        .Font.Bold = True
    End With

    Range("K" & emptyRow + 1).value = Tt1                  '\\ Uren totaal
    Range("M" & emptyRow + 1).value = Tt2                  '\\ Fooi totaal
    
    Columns("K:K").Select
    Selection.NumberFormat = "General"
    
    Range("A1").Select
    'TODO Error handling toevoegen
    '\\ Error handling toevoegen
    [1:200].value = [1:200].value

    Sheets("WEAKLYTEMP").Delete                            '\\ Comment verwijderen wanneer kan
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

    Unload Me

    Exit Sub

Error_Value:

    MsgBox "Er is een fout opgetreden, probeer het opnieuw aub" & vbNewLine & vbNewLine & _
           ". Fout kan zijn opgetreden door; #WAARDE! fout."

    Sheets("WEAKLYTEMP").Delete
    Sheets("INDEX").Select

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

    Unload Me
    
End Sub

Private Sub CmdCan_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    #If Mac Then
        ResizeUserForm Me
    #End If

    Dim X As Integer

    For X = 1 To 31                                        'populate comboBox with days
        Me.DD1.AddItem X
    Next

    For X = 1 To 12                                        'populate comboBox with Months
        Me.MM1.AddItem X
    Next

    For X = 2020 To 2030                                   'populate comboBox with Years
        Me.YYYY1.AddItem X
    Next

    Dim myDate As Date
    myDate = Date - 7

    DD1.value = Format(myDate, "dd")                       'Gets the current date in the comboBoxes
    MM1.value = Format(myDate, "mm")
    YYYY1.value = Format(myDate, "yyyy")
    
    '   DD1.Value = Left(myDate, 2)                  'Old way!
    '   MM1.Value = Mid(myDate, 4, 2)
    '   YYYY1.Value = Mid(myDate, 7, 4)
    '   YYYY1.ListIndex = 1

End Sub

