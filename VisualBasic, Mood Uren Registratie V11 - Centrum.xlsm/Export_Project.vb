Attribute VB_Name = "Export_Project"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Sub ExportProcess()
    Dim answer As Integer

    If Range("B15").value = "Mood Eindhoven, Willem Aarts" Then
        
        answer = MsgBox("Export Active project to GitHub?" & vbNewLine & vbNewLine & _
                        "Directory;" & vbNewLine & _
                        "C:\Users\wille\OneDrive\Documenten\GitHub\MOOD_Uren_Bestand\VisualBasic, " & ActiveWorkbook.Name, vbExclamation + vbYesNo)
    
        If answer = vbNo Then
            'MsgBox "Closing workbook"
            Exit Sub
        Else
            MsgBox "Exporting Active project to GitHub map"
            Application.Run ("Export_Project.ExportVisualBasicCode")
            'Code
        End If
    
    Else
        'Do Nothing
    End If
    
End Sub

Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    
    'directory = ActiveWorkbook.path & "\VisualBasic, " & ActiveWorkbook.Name
    directory = "C:\Users\wille\OneDrive\Documenten\GitHub\MOOD_Uren_Bestand\VisualBasic, " & ActiveWorkbook.Name
    
    count = 0
    
    If Dir(directory, vbDirectory) = "" Then
        MkDir directory
    End If
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
        Case ClassModule, Document
            extension = ".vb"                              '".cls"
        Case Form
            extension = ".vb"                              '".frm"
        Case Module
            extension = ".vb"                              '".bas"
        Case Else
            GoTo NtN
            'extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
NtN:
    Next
    
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    Application.StatusBar = False
End Sub


