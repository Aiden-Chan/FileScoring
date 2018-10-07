Attribute VB_Name = "B_GetFiles"
Option Explicit

Sub FileScoring()
    FSForm.Show
End Sub


Sub FindFiles()
    
    Dim FilePath As String
    Dim TestFileName As String
    Dim SourceRange As String
    Dim i As Integer, j As Integer
    Dim PrintRow As Integer
    Dim Symbols() As String
    
    NumFiles = 0
    Symbols = Split(" 0 1 2 3 4 5 6 7 8 9 ! "" £ $ % ^ & * ( ) - _ = + { [ } ] ; : @ ' ~ # < , > . ? / ")
    
    'Extract names of all text files of the correct type from the folder.
    TestFileName = Dir(FolderPath & "*" & FileType)
    While TestFileName <> ""
        NumFiles = NumFiles + 1
        ReDim Preserve NameFile(0 To NumFiles)
        NameFile(NumFiles) = Left(TestFileName, Len(TestFileName) - Len(FileType))
        
        TestFileName = Dir() 'Get the next file that matches the description
    Wend
    
    'Resize variables. Index 0 contains data from all text files
    ReDim ContentFile(0 To NumFiles)
    ReDim NumWords(0 To NumFiles)
    NameFile(0) = "Collated"
    
    For i = 1 To NumFiles
        'Extract text as a string
        Application.StatusBar = "Reading " & NameFile(i) & FileType
        FilePath = FolderPath & NameFile(i) & FileType
        Open FilePath For Input As #i
        ContentFile(i) = Input(LOF(i), i)
        Close #i
        
        'Remove symbols and line breaks
        For j = 0 To UBound(Symbols) - 1
            ContentFile(i) = Replace(ContentFile(i), Symbols(j), "", , , vbBinaryCompare)
        Next j
        ContentFile(i) = Replace(ContentFile(i), vbNewLine, " ", , , vbBinaryCompare)
        
        'Store all text in index 0
        ContentFile(0) = ContentFile(0) & " " & ContentFile(i)
    Next i
    
    For i = 0 To NumFiles
        If i > 0 Or NumFiles <> 1 Then 'Don't bother doing a Collated report if there is only one text file
            Application.StatusBar = "Analysing " & NameFile(i) & "..."
            
            'Split the one long string into individual words, separated by spaces
            Words = Split(ContentFile(i))
            Call NewSheet("Words_" & NameFile(i))
            
            'Print all words on a separate row, ignoring blanks
            PrintRow = 2
            Sheets("Words_" & NameFile(i)).Cells(1, 1) = "Words"
            Sheets("Words_" & NameFile(i)).Cells(1, 1).Font.Bold = True
            For j = 1 To UBound(Words)
                If Words(j - 1) <> "" Then
                    Sheets("Words_" & NameFile(i)).Cells(PrintRow, 1) = Words(j - 1)
                    PrintRow = PrintRow + 1
                End If
            Next j
            NumWords(i) = PrintRow - 1
            
            SourceRange = Sheets("Words_" & NameFile(i)).Range("A1:A" & PrintRow).Address
            
            'Make pivot tables
            Call NewSheet("Pivot_" & NameFile(i))
            Call MakePivot(i, SourceRange)
        End If
    Next i
    
    Application.StatusBar = False
End Sub

Function NewSheet(SheetName As String)
    Dim NumSheets As Integer
    
    'If the sheet name already exists, overwrite it
    'Otherwise, make a new sheet of that name
    'The Error lines allow for a non-existent sheet to be 'deleted'
    'The DisplayAlerts line disables the 'Are you sure you want to delete these sheets?' prompt
    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets(SheetName).Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
    
    NumSheets = ActiveWorkbook.Sheets.Count
    Sheets.Add After:=Sheets(NumSheets)
    Sheets(NumSheets + 1).Name = SheetName
    
End Function
