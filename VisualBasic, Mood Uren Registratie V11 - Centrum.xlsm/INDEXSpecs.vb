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
    
    Case "Menu5"
        Debug.Print "Menu5 - WeaklyClick"
        DeleteEmployee.Show

    Case "Login"
        Debug.Print "Login - LoginClick"
        Login.Show
        
    Case Else
        Debug.Print "None Selected"

    End Select

End Sub

Sub GoToINDEX()
    'Keeps all the important sheets, the sheets that are suppose to be deleted will be deleted

    Dim Shn As Variant
    Shn = ActiveSheet.Name

    Application.DisplayAlerts = False

    Select Case Shn

    Case "DataStr"
        Debug.Print "DataStr - Click"
        Sheets("INDEX").Select
    
    Case "DataEmp"
        Debug.Print "DataEmp - Click"
        Sheets("INDEX").Select

    Case "<EMP>"
        Debug.Print "<EMP> - Click"
        Sheets("INDEX").Select

    Case "TEMP-MTseven"
        Debug.Print "TEMP-MTseven - Click"
        Sheets("INDEX").Select

    Case "TEMP-TOTAL"
        Debug.Print "TEMP-TOTAL - Click"
        Sheets("INDEX").Select

    Case "INDEX"
        Debug.Print "INDEX - Click"
        Sheets("INDEX").Select
        
    Case "DataArchive"
        Debug.Print "DataArchive - Click"
        Sheets("INDEX").Select

    Case "TEMP-WEAKLY"
        Debug.Print "TEMP-WEAKLY - Click"
        Columns("A:M").ClearContents
        
        Application.ScreenUpdating = False
        Columns("A:M").Select
        With Selection
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
        End With
        Range("A1").Select
        Application.ScreenUpdating = True
        
        Sheets("INDEX").Select

    Case Else
        Debug.Print Shn & "; Deleted."
        ActiveWindow.SelectedSheets.Delete
        Sheets("INDEX").Select

    End Select

    Application.DisplayAlerts = False

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

