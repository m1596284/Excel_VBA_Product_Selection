Attribute VB_Name = "Module1"

Sub Create_Import()
    
    Application.DisplayAlerts = False

    'Create
'    Sheets.Add After:=ActiveSheet
'    Sheets(2).Select
'    Sheets(2).Name = "Cover Sheet"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(3).Select
    Sheets(3).Name = "Price Book"

    'Import
    Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    If fDialog.Show = -1 Then
    End If
    fieldArray = Split(fDialog.SelectedItems(1), "\")
    Workbooks.Open Filename:=fDialog.SelectedItems(1)
    Windows("總表_v11.xlsm").Activate
    Windows(fieldArray(UBound(fieldArray))).Activate
    'copy_paste_close
    Dim rowEnd As Long
    rowEnd = 0
    Sheets("INTL Price List").Select
    Cells(6, 3).Select
    rowEnd = Selection.End(xlDown).Row
    Range(Cells(6, 1), Cells(rowEnd, 7)).Select
    Selection.Copy
    Windows("總表_v11.xlsm").Activate
    Worksheets(3).Select
    Cells(1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    rowEnd = Cells(1, 1).End(xlDown).Row
    ActiveWorkbook.Worksheets(3).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(3).Sort.SortFields.Add2 Key:=Range("B1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(3).Sort
        .SetRange Range(Cells(2, 1), Cells(rowEnd, 7))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Windows(fieldArray(UBound(fieldArray))).Activate
    ActiveWindow.Close

    'Parameters_Load
    Sheets(1).Select
    Dim Gold As Variant
    Dim Silver As Variant
    Dim Gold2 As Single
    Dim Silver2 As Single
    Gold = Cells(1, 1)
    Silver = Cells(2, 1)
    Gold2 = Cells(1, 2)
    Silver2 = Cells(2, 2)
    Sheets(3).Select
    Cells(1, 9) = Gold
    Cells(2, 9) = Silver
    Cells(1, 10) = Gold2
    Cells(2, 10) = Silver2
    ActiveWorkbook.Names.Add Name:="Level", RefersToR1C1:="='Price Book'!R1C9:R2C9"
    ActiveWorkbook.Names.Add Name:="Level2", RefersToR1C1:="='Price Book'!R1C9:R2C10"
    Sheets(1).Select
    
    Dim UpArr() As Variant
    ReDim UpArr(Cells(5, 1).End(xlDown).Row - 6)
    Dim UpArr2() As Variant
    ReDim UpArr2(Cells(5, 1).End(xlDown).Row - 6)
    For i = 0 To UBound(UpArr)
        UpArr(i) = Cells(i + 6, 1)
    Next
    For i = 0 To UBound(UpArr2)
        UpArr2(i) = Cells(i + 6, 2)
    Next

    'Uplift !
    Dim stone As Long
    Dim upEnd As Long
    Sheets(3).Select
    upEnd = Cells(1, 2).End(xlDown).Row
    For i = 2 To upEnd
        If IsNumeric(Cells(i, 7)) Then
        Cells(i, 7) = Cells(i, 7) * UpArr2(Application.Match(Cells(i, 2), UpArr, 0) - 1)
        End If
    Next


    'From Here is the Category list
    'Data Processing
    Dim tempEnd As Long
    Dim mEnd As Variant
    Dim mArr() As Variant
    Dim moArr() As Variant
    Dim cArr() As Variant
    Dim CaArr() As Variant
    Dim y As Long
    Dim yEnd As Long
    Dim x As Long
    y = 1
    yEnd = 1
    x = 13
    Sheets(3).Select
    Cells(1, 4).Select
    tempEnd = Selection.End(xlDown).Row
    ActiveWorkbook.Names.Add Name:="List", RefersToR1C1:="='Price Book'!R2C5:R" & tempEnd & "C7"
    Range(Cells(1, 4), Cells(tempEnd, 4)).Select
    Selection.Copy
    Cells(1, 11).Select
    ActiveSheet.Paste
    ActiveSheet.Range(Cells(1, 11), Cells(tempEnd, 11)).RemoveDuplicates Columns:=1, Header:=xlYes
    Cells(1, 11).Select
    tempEnd = Selection.End(xlDown).Row
    
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Price Book").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Price Book").Sort.SortFields.Add2 Key:=Range("K2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Price Book").Sort
        .SetRange Range(Cells(2, 11), Cells(tempEnd, 11))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ReDim moArr(tempEnd - 2)
    For i = 2 To tempEnd
        moArr(i - 2) = UCase(Cells(i, 11))
    Next
    
    Cells(1, 11).Select
    tempEnd = Selection.End(xlDown).Row
    Range(Cells(1, 11), Cells(tempEnd, 11)).Select
    Selection.Replace What:="GLOBALPROTECT", Replacement:="GP", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:=":", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="\", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="[", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="]", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="'", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False


    'Name List
    Range(Cells(1, 11), Cells(tempEnd, 11)).Select
    Selection.Copy
    Cells(1, 12).Select
    ActiveSheet.Paste
    Selection.Replace What:=" ", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="-", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="\", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="'", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    For i = 1 To tempEnd
    If IsNumeric(Left(Cells(i, 12), 1)) Then
        Cells(i, 12) = "_" & Cells(i, 12)
    End If
    Next
    ReDim mArr(tempEnd - 2)
    For i = 2 To tempEnd
        If Len(Cells(i, 12)) > 31 Then
            mArr(i - 2) = Left(Cells(i, 12), 31)
        Else
            mArr(i - 2) = Cells(i, 12)
        End If
    Next
    Range("K1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Names.Add Name:="Model", RefersToR1C1:="='Price Book'!R2C11:R" & tempEnd & "C11"
    ActiveWorkbook.Names.Add Name:="Model2", RefersToR1C1:="='Price Book'!R2C11:R" & tempEnd & "C12"
    ActiveWorkbook.Names.Add Name:="Model3", RefersToR1C1:="='Price Book'!R2C12:R" & tempEnd & "C12"

    'Category
    Cells(1, 2).Select
    tempEnd = Selection.End(xlDown).Row
    Range(Cells(1, 2), Cells(tempEnd, 2)).Select
    Selection.Copy
    Cells(4, 9).Select
    ActiveSheet.Paste
    ActiveSheet.Range(Cells(4, 9), Cells(tempEnd + 3, 9)).RemoveDuplicates Columns:=1, Header:=xlYes
    Cells(4, 9).Select
    tempEnd = Selection.End(xlDown).Row
    ReDim CaArr(tempEnd - 5)
    For i = 5 To tempEnd
        CaArr(i - 5) = UCase(Cells(i, 9))
    Next
    
    Cells(4, 9).Select
    tempEnd = Selection.End(xlDown).Row
    Range(Cells(4, 9), Cells(tempEnd, 9)).Select
    Selection.Replace What:="GLOBALPROTECT", Replacement:="GP", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:=":", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="\", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="[", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="]", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="'", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False


    'Name List
    Range(Cells(4, 9), Cells(tempEnd, 9)).Select
    Selection.Copy
    Cells(4, 10).Select
    ActiveSheet.Paste
    Selection.Replace What:=" ", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="-", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="\", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="'", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    For i = 4 To tempEnd
    If IsNumeric(Left(Cells(i, 10), 1)) Then
        Cells(i, 10) = "_" & Cells(i, 10)
    End If
    Next
    ReDim cArr(tempEnd - 5)
    For i = 5 To tempEnd
        If Len(Cells(i, 10)) > 31 Then
            cArr(i - 5) = Left(Cells(i, 10), 31)
        Else
            cArr(i - 5) = Cells(i, 10)
        End If
    Next
    
    

    'Deform_Name it_Create sheets
    For i = 0 To UBound(cArr)
    Sheets(3).Select
    mEnd = CaArr(i)
    Do
        yEnd = yEnd + 1
        mEnd = UCase(Cells(yEnd, 2))
    Loop Until mEnd <> CaArr(i)
    If y = 1 Then
        y = 2
    End If
    Range(Cells(y, 5), Cells(yEnd - 1, 5)).Select
    Selection.Copy
    Cells(1, x) = cArr(i)
    Cells(2, x).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range(Cells(1, x), Cells(Cells(1, x).End(xlDown).Row, x)).Select
    ActiveWorkbook.Worksheets(3).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(3).Sort.SortFields.Add2 Key:=Cells(1, x) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(3).Sort
        .SetRange Range(Cells(2, x), Cells(Cells(2, x).End(xlDown).Row, x))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False
    Range(Cells(y, 1), Cells(yEnd - 1, 7)).Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = cArr(i)
    Sheets(Sheets.Count).Select
    Cells(2, 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Cells(1, 2) = "Group"
    Cells(1, 3) = "Category"
    Cells(1, 4) = "Product"
    Cells(1, 5) = "Model"
    Cells(1, 6) = "Part Name"
    Cells(1, 7) = "Description"
    Cells(1, 8) = "List Price"
    Cells(1, 1) = "Back"
    Cells(1, 1).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'Cover Sheet'!A1", TextToDisplay:="Back"
    
    'Auto Fit
    Cells.Select
    Cells.EntireColumn.AutoFit
    'Freeze Top
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    For ii = 2 To Cells(2, 2).End(xlDown).Row
        If UCase(Cells(ii, 2)) = "PRODUCT" Then
            If UCase(Cells(ii, 4)) <> "PLATFORMS" Then
                Cells(ii, 2) = "Product2"
            End If
        End If
    Next
    
    
    Columns("B:B").Select
    Application.AddCustomList ListArray:=Array("Product", "Support", "Subscription", "Product2")
    ActiveWorkbook.Worksheets(Sheets.Count).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(Sheets.Count).Sort.SortFields.Add2 Key:=Range(Cells(2, 2), Cells(Cells(2, 2).End(xlDown).Row, 2)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "Product,Support,Subscription,Product2", DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(Sheets.Count).Sort
        .SetRange Range(Cells(1, 2), Cells(Cells(1, 2).End(xlDown).Row, 7))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    For ii = 2 To Cells(1, 2).End(xlDown).Row
        If UCase(Cells(ii, 2)) = "PRODUCT2" Then
            Cells(ii, 2) = "Product"
        End If
    Next
    
    
    
    'Add Line
    Dim yy As Long
    Dim yy2 As Long
    Dim yyEnd As Long
    Dim GroupList As Integer
    Dim GroupTop As Integer
    Dim GroupEnd As Integer
    GroupList = 11
    yy = 2
    yy2 = 2
    yyEnd = Cells(1, 2).End(xlDown).Row + 1
    'Color White
    Range("J1:M1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    'Platforms
    If UCase(Cells(2, 4)) = "PLATFORMS" Then
        Cells(1, GroupList) = "Platforms"
        GroupList = GroupList + 1
        For ii = yy To yyEnd
            If UCase(Cells(ii, 4)) = "PLATFORMS" Then
                yy2 = Cells(ii, 4).Row
            Else
                Exit For
            End If
        Next
        'Add Name
        ActiveWorkbook.Names.Add Name:=cArr(i) & "Platforms", RefersToR1C1:="='" & cArr(i) & "'!R" & yy & "C6:R" & yy2 & "C6"
        'Add row
        Rows(2).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        'Color
        Range(Cells(yy, 2), Cells(yy, 8)).Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        ActiveCell.FormulaR1C1 = "Product - Platforms"
        yy = yy2 + 2
        yy2 = yy2 + 2
        yyEnd = yyEnd + 1
    Else
    End If
    
    'Support
    If UCase(Cells(yy, 2)) = "SUPPORT" Then
        Cells(1, GroupList) = "Support"
        GroupList = GroupList + 1
        For ii = yy To yyEnd
            If UCase(Cells(ii, 2)) = "SUPPORT" Then
                yy2 = Cells(ii, 2).Row
            Else
                Exit For
            End If
        Next
        'Add Name
        ActiveWorkbook.Names.Add Name:=cArr(i) & "Support", RefersToR1C1:="='" & cArr(i) & "'!R" & yy & "C6:R" & yy2 & "C6"
        'Color
        Range(Cells(yy, 2), Cells(yy, 8)).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        ActiveCell.FormulaR1C1 = "Support"
        yy = yy2 + 2
        yy2 = yy2 + 2
        yyEnd = yyEnd + 1
    Else
    End If
    
    'Subscription
    If UCase(Cells(yy, 2)) = "SUBSCRIPTION" Then
        Cells(1, GroupList) = "Subscription"
        GroupList = GroupList + 1
        For ii = yy To yyEnd
            If UCase(Cells(ii, 2)) = "SUBSCRIPTION" Then
                yy2 = Cells(ii, 2).Row
            Else
                Exit For
            End If
        Next
        'Add Name
        ActiveWorkbook.Names.Add Name:=cArr(i) & "Subscription", RefersToR1C1:="='" & cArr(i) & "'!R" & yy & "C6:R" & yy2 & "C6"
        'Color
        Range(Cells(yy, 2), Cells(yy, 8)).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        ActiveCell.FormulaR1C1 = "Subscription"
        yy = yy2 + 2
        yy2 = yy2 + 2
        yyEnd = yyEnd + 1
    Else
    End If
    
    'Accessories
    If UCase(Cells(yy, 2)) = "PRODUCT - PLATFORMS" Then
    ElseIf UCase(Cells(yy, 2)) = "SUPPORT" Then
    ElseIf UCase(Cells(yy, 2)) = "SUBSCRIPTION" Then
    ElseIf Cells(yy, 2) = "" Then
    Else
        Cells(1, GroupList) = "Accessories"
        'Add Name
        ActiveWorkbook.Names.Add Name:=cArr(i) & "Accessories", RefersToR1C1:="='" & cArr(i) & "'!R" & yy & "C6:R" & yy2 & "C6"
        Range(Cells(yy, 2), Cells(yy, 8)).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        ActiveCell.FormulaR1C1 = "Accessories"
    End If
    
    GroupTop = Cells(1, 10).End(xlToRight).Column
    GroupEnd = Cells(1, 20).End(xlToLeft).Column
    ActiveWorkbook.Names.Add Name:=cArr(i) & "Group", RefersToR1C1:="='" & cArr(i) & "'!R1C" & GroupTop & ":R1C" & GroupEnd & ""
    
    Cells(2, 1).Select
    y = yEnd
    x = x + 1
    Next
    
    ''From here is model type
    y = 1
    yEnd = 1
    Sheets(3).Select
    Cells(1, 4).Select
    tempEnd = Selection.End(xlDown).Row
    Columns("D:D").Select
    ActiveWorkbook.Worksheets("Price Book").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Price Book").Sort.SortFields.Add2 Key:=Range(Cells(2, 4), Cells(tempEnd, 4)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Price Book").Sort
        .SetRange Range(Cells(1, 1), Cells(tempEnd, 7))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
        
    
    'Model
    'Deform_Name it_Create sheets
    For i = 0 To UBound(mArr)
    Sheets(3).Select
    mEnd = moArr(i)
    Do
        yEnd = yEnd + 1
        mEnd = UCase(Cells(yEnd, 4))
    Loop Until mEnd <> moArr(i)
    If y = 1 Then
        y = 2
    End If
    Range(Cells(y, 5), Cells(yEnd - 1, 5)).Select
    Selection.Copy
    Cells(1, x) = mArr(i)
    Cells(2, x).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range(Cells(1, x), Cells(Cells(1, x).End(xlDown).Row, x)).Select
    ActiveWorkbook.Worksheets(3).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(3).Sort.SortFields.Add2 Key:=Cells(1, x) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(3).Sort
        .SetRange Range(Cells(2, x), Cells(Cells(2, x).End(xlDown).Row, x))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False
    Range(Cells(y, 1), Cells(yEnd - 1, 7)).Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = mArr(i)
    Sheets(Sheets.Count).Select
    Cells(2, 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Cells(1, 2) = "Group"
    Cells(1, 3) = "Category"
    Cells(1, 4) = "Product"
    Cells(1, 5) = "Model"
    Cells(1, 6) = "Part Name"
    Cells(1, 7) = "Description"
    Cells(1, 8) = "List Price"
    Cells(1, 1) = "Back"
    Cells(1, 1).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'Cover Sheet'!A1", TextToDisplay:="Back"
    
    'Auto Fit
    Cells.Select
    Cells.EntireColumn.AutoFit
    'Freeze Top
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    For ii = 2 To Cells(2, 2).End(xlDown).Row
        If UCase(Cells(ii, 2)) = "PRODUCT" Then
            If UCase(Cells(ii, 4)) <> "PLATFORMS" Then
                Cells(ii, 2) = "Product2"
            End If
        End If
    Next
    
    
    Columns("B:B").Select
    Application.AddCustomList ListArray:=Array("Product", "Support", "Subscription", "Product2")
    ActiveWorkbook.Worksheets(Sheets.Count).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(Sheets.Count).Sort.SortFields.Add2 Key:=Range(Cells(2, 2), Cells(Cells(2, 2).End(xlDown).Row, 2)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "Product,Support,Subscription,Product2", DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(Sheets.Count).Sort
        .SetRange Range(Cells(1, 2), Cells(Cells(1, 2).End(xlDown).Row, 7))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    For ii = 2 To Cells(1, 2).End(xlDown).Row
        If UCase(Cells(ii, 2)) = "PRODUCT2" Then
            Cells(ii, 2) = "Product"
        End If
    Next
    
    
    
    'Add Line
'    Dim yy As Long
'    Dim yy2 As Long
'    Dim yyEnd As Long
'    Dim GroupList As Integer
'    Dim GroupTop As Integer
'    Dim GroupEnd As Integer
    GroupList = 11
    yy = 2
    yy2 = 2
    yyEnd = Cells(1, 2).End(xlDown).Row + 1
    'Color White
    Range("J1:M1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    'Platforms
    If UCase(Cells(2, 4)) = "PLATFORMS" Then
        Cells(1, GroupList) = "Platforms"
        GroupList = GroupList + 1
        For ii = yy To yyEnd
            If UCase(Cells(ii, 4)) = "PLATFORMS" Then
                yy2 = Cells(ii, 4).Row
            Else
                Exit For
            End If
        Next
        'Add Name
        ActiveWorkbook.Names.Add Name:=mArr(i) & "Platforms", RefersToR1C1:="='" & mArr(i) & "'!R" & yy & "C6:R" & yy2 & "C6"
        'Add row
        Rows(2).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        'Color
        Range(Cells(yy, 2), Cells(yy, 8)).Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        ActiveCell.FormulaR1C1 = "Product - Platforms"
        yy = yy2 + 2
        yy2 = yy2 + 2
        yyEnd = yyEnd + 1
    Else
    End If
    
    'Support
    If UCase(Cells(yy, 2)) = "SUPPORT" Then
        Cells(1, GroupList) = "Support"
        GroupList = GroupList + 1
        For ii = yy To yyEnd
            If UCase(Cells(ii, 2)) = "SUPPORT" Then
                yy2 = Cells(ii, 2).Row
            Else
                Exit For
            End If
        Next
        'Add Name
        ActiveWorkbook.Names.Add Name:=mArr(i) & "Support", RefersToR1C1:="='" & mArr(i) & "'!R" & yy & "C6:R" & yy2 & "C6"
        'Color
        Range(Cells(yy, 2), Cells(yy, 8)).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        ActiveCell.FormulaR1C1 = "Support"
        yy = yy2 + 2
        yy2 = yy2 + 2
        yyEnd = yyEnd + 1
    Else
    End If
    
    'Subscription
    If UCase(Cells(yy, 2)) = "SUBSCRIPTION" Then
        Cells(1, GroupList) = "Subscription"
        GroupList = GroupList + 1
        For ii = yy To yyEnd
            If UCase(Cells(ii, 2)) = "SUBSCRIPTION" Then
                yy2 = Cells(ii, 2).Row
            Else
                Exit For
            End If
        Next
        'Add Name
        ActiveWorkbook.Names.Add Name:=mArr(i) & "Subscription", RefersToR1C1:="='" & mArr(i) & "'!R" & yy & "C6:R" & yy2 & "C6"
        'Color
        Range(Cells(yy, 2), Cells(yy, 8)).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        ActiveCell.FormulaR1C1 = "Subscription"
        yy = yy2 + 2
        yy2 = yy2 + 2
        yyEnd = yyEnd + 1
    Else
    End If
    
    'Accessories
    If UCase(Cells(yy, 2)) = "PRODUCT - PLATFORMS" Then
    ElseIf UCase(Cells(yy, 2)) = "SUPPORT" Then
    ElseIf UCase(Cells(yy, 2)) = "SUBSCRIPTION" Then
    ElseIf Cells(yy, 2) = "" Then
    Else
        Cells(1, GroupList) = "Accessories"
        'Add Name
        ActiveWorkbook.Names.Add Name:=mArr(i) & "Accessories", RefersToR1C1:="='" & mArr(i) & "'!R" & yy & "C6:R" & yy2 & "C6"
        Range(Cells(yy, 2), Cells(yy, 8)).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        ActiveCell.FormulaR1C1 = "Accessories"
    End If
    
    GroupTop = Cells(1, 10).End(xlToRight).Column
    GroupEnd = Cells(1, 20).End(xlToLeft).Column
    ActiveWorkbook.Names.Add Name:=mArr(i) & "Group", RefersToR1C1:="='" & mArr(i) & "'!R1C" & GroupTop & ":R1C" & GroupEnd & ""
    
    Cells(2, 1).Select
    y = yEnd
    x = x + 1
    ActiveWindow.SelectedSheets.Visible = False
    
    Next
    
    

    'Cover Sheet
    Sheets(2).Select
    'Draw range
    Range("C2:H21").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("I1").Select
    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Parameter!$A$1:$A$2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Cells(1, 8) = "Level"
    Cells(1, 9) = Gold
    Cells(2, 3) = "Model"
    Cells(2, 4) = "Group"
    Cells(2, 5) = "Part Name"
    Cells(2, 6) = "Number"
    Cells(2, 7) = "List Price"
    Cells(2, 8) = "Discount"
    Cells(2, 9) = "Amount"
    Cells(2, 10) = "PN# description"
    Cells(21, 8) = "Total"
    'Level
    Range("I1").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Level"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
'    'Amount
'    Range("H3").Select
'    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-3]*(RC[-2]-RC[-1]),"""")"
'    Selection.AutoFill Destination:=Range("H3:H20"), Type:=xlFillDefault
'    'Sum
'    Range("H21").Select
'    ActiveCell.FormulaR1C1 = "=SUM(R[-18]C:R[-1]C)"
'    'List Price
'    Range("F3").Select
'    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],List,3,0),"""")"
'    Selection.AutoFill Destination:=Range("F3:F20"), Type:=xlFillDefault
'    'DisCount
'    Range("G3").Select
'    ActiveCell.FormulaR1C1 = "=IFERROR((1-VLOOKUP(R1C7,Level2,2,0))*'Cover Sheet'!RC[-1],"""")"
'    Selection.AutoFill Destination:=Range("G3:G20")
'    'Model
'    Range("C3:C20").Select
'    With Selection.Validation
'        .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:="=OFFSET(Model3,MATCH(" * "&C3&" * ",Model3,0)-1,0,COUNTIF(Model3," * "&C3&" * "),1)"
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = ""
'        .InputMessage = ""
'        .ErrorMessage = ""
'        .IMEMode = xlIMEModeNoControl
'        .ShowInput = True
'        .ShowError = True
'    End With
'    Range("D3:D20").Select
'    With Selection.Validation
'        .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:="=INDIRECT(C3)"
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = ""
'        .InputMessage = ""
'        .ErrorMessage = ""
'        .IMEMode = xlIMEModeNoControl
'        .ShowInput = True
'        .ShowError = True
'    End With

    'Category
    Cells(2, 1) = "Category"
    For i = 0 To UBound(cArr)
        Cells(i + 3, 1) = cArr(i)
        Cells(i + 3, 1).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & cArr(i) & "'!A1", TextToDisplay:=cArr(i)
    Next
    Columns("A:A").EntireColumn.AutoFit
    
    'Delete Parameter
    Sheets("Parameter").Select
    Workbooks("總表_v11.xlsm").Sheets("Parameter").Delete

    'Locker
    Sheets("Cover Sheet").Select
    ActiveSheet.Unprotect
    Cells.Select
    Selection.Locked = True
    Selection.FormulaHidden = True
    Range("C3:F20,A:A,I1").Select
    Selection.Locked = False
    Selection.FormulaHidden = True
    Sheets("Cover Sheet").Select
    ActiveSheet.Protect Password:="123123123", DrawingObjects:=True, Contents:=True, Scenarios:=True
    Sheets("Price Book").Select
    ActiveWindow.SelectedSheets.Visible = False
    ActiveWorkbook.Protect Password:="123123123", Structure:=True, Windows:=False
    Sheets(1).Select

    'Export_Save
    strNowPath = Excel.ActiveWorkbook.Path
    ChDir strNowPath
    ActiveWorkbook.SaveAs Filename:=strNowPath & "\Pricebook_Update.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    Cells(3, 3).Select
    
    
    
    Application.DisplayAlerts = True

End Sub




