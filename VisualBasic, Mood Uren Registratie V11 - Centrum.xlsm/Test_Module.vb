Attribute VB_Name = "Test_Module"
Private Sub Test_EmptyRow_HideEmployee()
    Dim emptyRow As Long
    Dim X As Integer

    emptyRow = WorksheetFunction.CountA(Range("B:B")) + 1

    Range("B1").Select

    Application.ScreenUpdating = False

    For X = 1 To emptyRow
        ActiveCell.Offset(1, 0).Select
        
        X = X + 1

        If ActiveCell.Offset(0, 9).value = CVErr(xlErrValue) Then
            Debug.Print "Error; #WAARDE!"
            MsgBox "A #WAARDE! error has occurred. Please contact the administrator"
        End If
        
        If ActiveCell.Offset(0, 9).value = "" Then
            Rows(X).EntireRow.Hidden = True
        End If
        
        X = X - 1
    Next

    Application.ScreenUpdating = True

End Sub

Private Sub Test_WAARDE_Error()
    If ActiveCell.value = CVErr(xlErrValue) Then
        Debug.Print "Error; #WAARDE!"
        MsgBox "A #WAARDE! error has occurred. Please contact the administrator"
    End If
End Sub

