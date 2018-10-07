Attribute VB_Name = "C_AnalyseText"
Option Explicit

Function MakePivot(i As Integer, SourceRange As String)
    Dim PvtShtName As String, PivotName As String
    Dim PivItem As PivotItem
    Dim SumTable As Range, PivotCorner As Range
    Dim Rank As Integer, RankRow As Integer
    Dim FreqWords As String 'Long string containing all frequent words
    
    PvtShtName = "Pivot_" & NameFile(i)
    PivotName = "PT_" & NameFile(i)
    
    Set SumTable = Range("A1")
    SumTable.Columns(1).ColumnWidth = 11.57
    Columns(1).HorizontalAlignment = xlLeft
    Set PivotCorner = SumTable(4, 2)
    SumTable(1, 1) = "File:"
    SumTable(1, 2) = NameFile(i)
    If i > 0 Then SumTable(1, 2) = SumTable(1, 2) & FileType
    SumTable(2, 1) = "Word count:"
    SumTable(2, 2) = NumWords(i)
    SumTable(2, 2).HorizontalAlignment = xlLeft

    'Make pivot table
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "'Words_" & NameFile(i) & "'!" & SourceRange, Version:=6).CreatePivotTable TableDestination:= _
        "'" & PvtShtName & "'!" & PivotCorner.Address(, , xlR1C1), TableName:=PivotName, DefaultVersion:=6
    PivotCorner.Select
    With Sheets(PvtShtName).PivotTables(PivotName)
        .ColumnGrand = False
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
        .RepeatAllLabels xlRepeatLabels
        .AddDataField Sheets(PvtShtName).PivotTables(PivotName).PivotFields("Words"), "Count of Words", xlCount
    End With
    With Sheets(PvtShtName).PivotTables(PivotName).PivotFields("Words")
        .Orientation = xlRowField
        .Position = 1
        .AutoSort xlDescending, "Count of Words", Sheets(PvtShtName).PivotTables(PivotName).PivotColumnAxis.PivotLines(1), 1
    End With
    PivotCorner = "Word"
    
    PivotCorner(1, 0) = "Rank"
    PivotCorner.Copy
    PivotCorner(1, 0).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    'Assign rank to words
    Rank = 1
    RankRow = 3
    PivotCorner(2, 0) = 1
    FreqWords = " " & PivotCorner(2, 1) & " "
    While Rank <= NumFreq And PivotCorner(RankRow, 1) <> "" And RankRow < Sheets(PvtShtName).PivotTables(PivotName).PivotFields("Words").PivotItems.Count - 2
        If PivotCorner(RankRow, 2) < PivotCorner(RankRow - 1, 2) Then Rank = RankRow - 1
        If Rank <= NumFreq Then
            PivotCorner(RankRow, 0) = Rank
            FreqWords = FreqWords & PivotCorner(RankRow, 1) & " "
            RankRow = RankRow + 1
        End If
    Wend
            
    'Hide words not in the top 10
    For Each PivItem In Sheets(PvtShtName).PivotTables(PivotName).PivotFields("Words").PivotItems
        If Len(Replace(FreqWords, " " & PivItem & " ", "")) = Len(FreqWords) Then PivItem.Visible = False
    Next PivItem
    
    'Count words not in the top 10
    PivotCorner(RankRow, 1) = "[Other]"
    PivotCorner(RankRow, 2) = "=" & SumTable(2, 2).Address & "-SUM(" & PivotCorner(2, 2).Address & ":" & PivotCorner(RankRow - 1, 2).Address & ")"
    
    'Export results to PDF
    If PDF Then
        Sheets(PvtShtName).ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        FolderPath & NameFile(i) & ".pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    End If
    
End Function

Sub DeleteSheet()
Attribute DeleteSheet.VB_ProcData.VB_Invoke_Func = "q\n14"
    Application.DisplayAlerts = False
    Sheets(ActiveWorkbook.Sheets.Count).Delete
    Application.DisplayAlerts = True
End Sub
