Attribute VB_Name = "M_UserFormSize"
Option Explicit

'@Ignore HungarianNotation
Public Const gUserFormResizeFactor As Double = 1.333333

'@Ignore ParameterCanBeByVal, ImplicitByRefModifier
Sub ResizeUserForm(Frm As Object, Optional dResizeFactor As Double = 0#)
    Dim ctrl As Control
    '@Ignore HungarianNotation
    Dim sColWidths As String
    Dim vColWidths As Variant
    '@Ignore HungarianNotation
    Dim iCol As Long

    If dResizeFactor = 0 Then dResizeFactor = gUserFormResizeFactor
    With Frm
        .Height = .Height * dResizeFactor
        .Width = .Width * dResizeFactor

        For Each ctrl In Frm.Controls
            With ctrl
                .Height = .Height * dResizeFactor
                .Width = .Width * dResizeFactor
                .Left = .Left * dResizeFactor
                .Top = .Top * dResizeFactor
                On Error Resume Next
                '@Ignore MemberNotOnInterface
                .Font.Size = .Font.Size * dResizeFactor
                On Error GoTo 0

                ' multi column listboxes, comboboxes
                Select Case TypeName(ctrl)
                Case "ListBox", "ComboBox"
                    '@Ignore MemberNotOnInterface
                    If ctrl.ColumnCount > 1 Then
                        '@Ignore MemberNotOnInterface
                        sColWidths = ctrl.ColumnWidths
                        vColWidths = Split(sColWidths, ";")
                        For iCol = LBound(vColWidths) To UBound(vColWidths)
                            vColWidths(iCol) = Val(vColWidths(iCol)) * dResizeFactor
                        Next
                        sColWidths = Join(vColWidths, ";")
                        '@Ignore MemberNotOnInterface
                        ctrl.ColumnWidths = sColWidths
                    End If
                End Select
            End With
        Next
    End With
End Sub

