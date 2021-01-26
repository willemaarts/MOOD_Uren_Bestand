Attribute VB_Name = "PDFmodule"
Option Explicit

Public Sub pdfWeakly()
    'uiteindelijk als je weg gaat uit sheet, alles verwijderen
    'Ron de Bruin : 29-July-2017
    'Test macro to save the ActiveSheet as pdf with ExportAsFixedFormat
    
    Dim FileName As String
    Dim FolderName As String
    Dim Folderstring As String
    Dim FilePathName As String
    Dim Date_1 As String
    Dim FacName As String

    #If Mac Then
        'If my ActiveSheet is landscape, I must attach this line
        'for making the PDF also landscape, seems to default to xlPortait
        ActiveSheet.PageSetup.Orientation = ActiveSheet.PageSetup.Orientation

        'Name of the folder in the Office folder
        FolderName = "PDFSaveFolder"
        'Name of the pdf file
        'Date_1 = Right(Range("C14").value, 5)
        FileName = "Wekelijkse urenstaat " & Range("B1").value & ".pdf"

        Folderstring = CreateFolderinMacOffice2016(NameFolder:=FolderName)
        FilePathName = Folderstring & Application.PathSeparator & FileName

        'expression A variable that represents a Workbook, Sheet, Chart, or Range object.
        'the parameters are not working like in Excel for Windows
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
                                        FilePathName, Quality:=xlQualityStandard, _
                                        IncludeDocProperties:=True, IgnorePrintAreas:=False
    
        MsgBox "PDF file saved in this location : " & FilePathName
    #Else
        ' Windows pdf macro
        FacName = Right(Range("B1").value, 4)
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
                                        ThisWorkbook.path & "\" & "Wekenlijkse urenstaat" & " " & FacName & ".pdf", _
                                        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                                        IgnorePrintAreas:=False, OpenAfterPublish:=True
        
    #End If
End Sub

Public Sub pdfEmployee()                                    'SaveActiveSheetAsPDFInMacExcel2016()
    'Ron de Bruin : 29-July-2017
    'Test macro to save the ActiveSheet as pdf with ExportAsFixedFormat
    
    Dim FileName As String
    Dim FolderName As String
    Dim Folderstring As String
    Dim FilePathName As String
    Dim Date_1 As String
    Dim FacName As String

    #If Mac Then
        'If my ActiveSheet is landscape, I must attach this line
        'for making the PDF also landscape, seems to default to xlPortait
        ActiveSheet.PageSetup.Orientation = ActiveSheet.PageSetup.Orientation

        'Name of the folder in the Office folder
        FolderName = "PDFSaveFolder"
        'Name of the pdf file
        Date_1 = Right(Range("C14").value, 5)
        FileName = ActiveSheet.Name & " " & Right(Range("C14").value, 5) & ".pdf"

        Folderstring = CreateFolderinMacOffice2016(NameFolder:=FolderName)
        FilePathName = Folderstring & Application.PathSeparator & FileName

        'expression A variable that represents a Workbook, Sheet, Chart, or Range object.
        'the parameters are not working like in Excel for Windows
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
                                        FilePathName, Quality:=xlQualityStandard, _
                                        IncludeDocProperties:=True, IgnorePrintAreas:=False
    
        MsgBox "PDF file saved in this location : " & FilePathName
    #Else
        ' Windows pdf macro
        FacName = Right(Range("C14").value, 4)
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
                                        ThisWorkbook.path & "\" & ActiveSheet.Name & " " & FacName & ".pdf", _
                                        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                                        IgnorePrintAreas:=False, OpenAfterPublish:=True
    #End If

End Sub

Function CreateFolderinMacOffice2016(NameFolder As String) As String
    'Function to create folder if it not exists in the Microsoft Office Folder
    'Ron de Bruin : 1-Feb-2019
    Dim OfficeFolder As String
    Dim PathToFolder As String
    Dim TestStr As String

    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") ' & _
                                                         ' "Library/Group Containers/UBF8T346G9.Office/"

    PathToFolder = OfficeFolder & NameFolder

    On Error Resume Next
    TestStr = Dir(PathToFolder & "*", vbDirectory)
    On Error GoTo 0
    If TestStr = vbNullString Then
        MkDir PathToFolder
        'You can use this msgbox line for testing if you want
        'MsgBox "You find the new folder in this location :" & PathToFolder
    End If
    CreateFolderinMacOffice2016 = PathToFolder
End Function



