Attribute VB_Name = "INDEXSpecs"
Option Explicit

Private Sub MenuOnClick()
    Dim Shp As Variant
    Shp = ActiveSheet.Shapes(Application.Caller).Name

    Select Case Shp
    
    Case "Menu1"
        Debug.Print "Menu1 - Employee click"
        EmployeeForm.Show

    Case "Menu2"
        Debug.Print "Menu2 - HourInpClick"
        HourInpForm.Show

    Case "Menu3"
        Debug.Print "Menu3 - NewEFormClick"
        NewEForm.Show

    Case "Menu4"
        Debug.Print "Menu4 - WeaklyClick"
        WeaklyForm.Show

    Case "Login"
        Debug.Print "Login - LoginClick"
        Login.Show
        
    Case Else
        Debug.Print "None Selected"

    End Select

End Sub

Sub GoToINDEX()
    'Keeps all the important sheets, the sheets that are suppose to be deleted will be deleted
    
    Application.DisplayAlerts = False

    If ActiveSheet.Name = "DataStr" Then
        Sheets("INDEX").Select
    ElseIf ActiveSheet.Name = "DataEmp" Then
        Sheets("INDEX").Select
    ElseIf ActiveSheet.Name = "<EMP>" Then
        Sheets("INDEX").Select
    ElseIf ActiveSheet.Name = "TEMP-MTseven" Then
        Sheets("INDEX").Select
    ElseIf ActiveSheet.Name = "TEMP-TOTAL" Then
        Sheets("INDEX").Select
    ElseIf ActiveSheet.Name = "INDEX" Then
        Sheets("INDEX").Select
    ElseIf ActiveSheet.Name = "TEMP-WEAKLY" Then
        Columns("A:M").ClearContents
        Sheets("INDEX").Select
    Else
        ActiveWindow.SelectedSheets.Delete
        Sheets("INDEX").Select
    End If

    Application.DisplayAlerts = True

End Sub

Private Sub Send_To_Admin()
    Dim myDate As Date
    myDate = Now()
    
    Dim answer As Integer
    answer = MsgBox("Correct closure for administrator?", vbQuestion + vbYesNo + vbDefaultButton2, "Close file")
    
    If answer = vbNo Then
        Exit Sub
    End If

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    On Error Resume Next
    Sheets("ADMIN").Select

    Range("B13").value = Range("B7").value
    Range("B15").value = myDate
    
    Range("B17").value = Range("B23").value
    Range("B23").value = ""
    
    Sheets("INDEX").Select
    
    ActiveWorkbook.Save
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
    MsgBox "You can now close the workbook"

End Sub

Sub ADMIN()
    'CNTR + SHIFT + Q
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = False
    End With

End Sub

Private Sub makenewyearInDataBase()

    Dim L1 As Date
    Dim L2 As Date
    Dim X As Integer
    
    L1 = "01-01-2021"

    ActiveCell.value = L1

    L2 = L1 + 1

    For X = 1 To 365
        ActiveCell.Offset(0, 4).Select
        ActiveCell.value = L2
        L2 = L2 + 1
    Next

End Sub

'Private Sub LoginClick()
'    'Shows the login screen
'    Login.Show
'End Sub
'
'Private Sub EmployeeClick()
'    'Shows the Employee screen
'    EmployeeForm.Show
'End Sub
'
'Private Sub HourInpClick()
'    'Shows the Hour Input screen
'    HourInpForm.Show
'End Sub
'
'Private Sub NewEFormClick()
'    'Shows the new employee screen
'    NewEForm.Show
'End Sub
'
'Private Sub WeaklyClick()
'    'Shows the Weakly screen
'    WeaklyForm.Show
'End Sub
    
Private Sub BasicCalendar()

    Dim dateVariable As Variant

    On Error Resume Next
    dateVariable = CalendarForm.GetDate( _
    OkayButton:=True, _
    ShowWeekNumbers:=True)
    
    Debug.Print dateVariable
    'If dateVariable <> 0 Then Range("A4") = dateVariable
End Sub


Private Sub AdvancedCalendar()
    dateVariable = CalendarForm.GetDate( _
        SelectedDate:=Range("H34").value, _
        FirstDayOfWeek:=Monday, _
        DateFontSize:=12, _
        TodayButton:=True, _
        OkayButton:=True, _
        ShowWeekNumbers:=True, _
        BackgroundColor:=RGB(243, 249, 251), _
        HeaderColor:=RGB(147, 205, 2221), _
        HeaderFontColor:=RGB(255, 255, 255), _
        SubHeaderColor:=RGB(223, 240, 245), _
        SubHeaderFontColor:=RGB(31, 78, 120), _
        DateColor:=RGB(243, 249, 251), _
        DateFontColor:=RGB(31, 78, 120), _
        TrailingMonthFontColor:=RGB(155, 194, 230), _
        DateHoverColor:=RGB(223, 240, 245), _
        DateSelectedColor:=RGB(202, 223, 242), _
        SaturdayFontColor:=RGB(0, 176, 240), _
        SundayFontColor:=RGB(0, 176, 240), _
        TodayFontColor:=RGB(0, 176, 80))
    If dateVariable <> 0 Then Range("H34") = dateVariable
End Sub
