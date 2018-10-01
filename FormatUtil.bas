 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class FormatUtil
'                                          to implement the format uniformly
'@author                                   Qiou Yang
'@lastUpdate                               01.10.2018
'                                          add bookmarks
'@TODO                                     add getter and setter
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declaration compulsory
Option Explicit

Dim pFont As String
Dim pFontSize As Integer
Dim pFontSizePlus As Integer
Dim pFontSizePlusPlus As Integer
Dim pRowHeightCaption As Integer
Dim pRowHeightFooting As Integer
Dim pRowHeight As Integer
Dim pBlankRowsUnterCaption As Integer
Dim pBgDark As Long
Dim pBgLight As Long

Public Enum FormatTheme
    fmtBlue = 0
    fmtLightGreen = 1
    fmtOrange = 2
    fmtSkyBlue = 3
    fmtBlackWhite = 4
End Enum

Private Sub Class_Initialize()
    pFont = "BahnSchrift"
    pFontSize = 10
    pFontSizePlus = pFontSize + 1
    pFontSizePlusPlus = pFontSize + 4
    pRowHeightCaption = 48
    pRowHeightFooting = 35
    pRowHeight = 24
    pBlankRowsUnterCaption = 2
    
    pBgDark = 11892015
    pBgLight = 16247773
End Sub

Function setTheme(Optional ByVal theme As Integer = FormatTheme.fmtLightGreen)
    If theme = FormatTheme.fmtLightGreen Then
        pBgDark = 9359529
        pBgLight = 14348258
    ElseIf theme = FormatTheme.fmtBlue Then
        pBgDark = 11892015
        pBgLight = 16247773
    ElseIf theme = FormatTheme.fmtOrange Then
        pBgDark = 3243501
        pBgLight = 14083324
    ElseIf theme = FormatTheme.fmtSkyBlue Then
        pBgDark = 15773696
        pBgLight = 16247773
    ElseIf theme = FormatTheme.fmtBlackWhite Then
        pBgDark = 7434613
        pBgLight = 14277081
    Else
        Err.Raise 9970, , "ParameterException: Unknown Theme: " & theme
    End If

End Function


Function formatRng(Optional ByRef rng As Range, Optional hasHeading As Boolean = True, Optional hasFooting As Boolean = False, Optional keepOldFormat As Boolean = False, Optional bgDark As Long, Optional bgLight As Long) As Range
    
    With Application
        .ScreenUpdating = False
    End With
    
    If IsMissing(rng) Or TypeName(rng) = "Nothing" Then
        Set rng = Application.Intersect(ActiveSheet.UsedRange, Selection)
    End If
    
    If IsMissing(bgDark) Or bgDark = 0 Then
        bgDark = pBgDark
    End If
    
    If IsMissing(bgLight) Or bgLight = 0 Then
        bgLight = pBgLight
    End If
    
    Dim bgTransparent As Long
    bgTransparent = -4142
    
    Dim fontColorWhite As Long
    Dim fontColorBlack As Long
    
    fontColorWhite = 16777215
    fontColorBlack = 0
    
    Dim contentStartRow As Long
    Dim contentEndRow As Long
    
    With rng

        contentStartRow = 1
        contentEndRow = .Rows.Count
        
        If Not keepOldFormat Then
            .ClearFormats
        End If
        ActiveWindow.DisplayGridlines = False
        
        ' borders internal vertical
        .Borders(xlInsideVertical).Weight = xlThin
        
        ' Alignment
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

        ' Font
        .Font.Name = pFont
        .Font.Size = pFontSize
        .Font.Color = fontColorBlack
        
        ' Row Height
        .rowHeight = pRowHeight
        
        If hasHeading Then
            With .Rows(1)
                With .Font
                    .Size = pFontSizePlus
                    .Color = fontColorWhite
                    .Bold = True
                End With
                .Interior.Color = bgDark
                
                .rowHeight = pRowHeightCaption
                
                .Borders(xlInsideVertical).Color = fontColorWhite
                .Borders(xlEdgeBottom).Weight = xlThin
            End With
            
            contentStartRow = 2
        End If
        
        If hasFooting Then
            With .Rows(.Rows.Count)
                With .Font
                    .Size = pFontSizePlus
                    .Color = fontColorWhite
                End With
                .Interior.Color = bgDark
                
                .rowHeight = pRowHeightFooting
            End With
            
            contentEndRow = contentEndRow - 1
        End If
        
        Dim i
        For i = contentStartRow To contentEndRow
            .Rows(i).Interior.Color = IIf(i Mod 2 = 1, bgLight, bgTransparent)
            .Rows(i).Borders(xlEdgeBottom).Weight = xlHairline
        Next i
        
        
        If hasFooting Then
            With .Rows(.Rows.Count)
                With .Borders(xlEdgeTop)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .Color = fontColorWhite
                End With
                
                .Borders(xlInsideVertical).Color = fontColorWhite
            End With
        End If
        

        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
    End With
    
    
    With Application
        .ScreenUpdating = True
    End With
    
    Set formatRng = rng
    
End Function


' caption occupies the first two rows above the table
' @param category : up-left   can be for example  Recherche / Arbeitspapier / Neuberechnung / Interview
Function addCaption(Optional ByRef rng As Range, Optional ByVal category As String = "Arbeitspapier", Optional ByVal theme As String = "Demo", Optional ByVal index As String = "M3-2018/001", Optional ByVal author As String = "Qiou Yang", Optional bgColor As Long) As Range
    
    With Application
        .ScreenUpdating = False
    End With
    
     If IsMissing(rng) Or TypeName(rng) = "Nothing" Then
        Set rng = Application.Intersect(ActiveSheet.UsedRange, Selection)
    End If
    
     If IsMissing(bgColor) Or bgColor = 0 Then
        bgColor = pBgDark
    End If
    
    Dim fontColorWhite As Long
    fontColorWhite = 16777215

    rng.Worksheet.Cells(1, rng.Cells(1, 1).Column).Resize(pBlankRowsUnterCaption + 2, rng.Columns.Count).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    With rng.Worksheet.Cells(1, rng.Cells(1, 1).Column).Resize(2, rng.Columns.Count)
        .Interior.Color = bgColor
        .Font.Name = pFont
        .Font.Size = pFontSizePlus
        .Font.Color = fontColorWhite
        
        .rowHeight = pRowHeightFooting
        
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        
        
        With .Cells(1, 1)
            .Value = category
        End With
        
        With .Cells(2, 1)
           .Value = theme
           .Font.Bold = True
           .Font.Size = pFontSizePlusPlus
        End With
        
        With .Cells(1, rng.Columns.Count)
           .Value = index
           .HorizontalAlignment = xlRight
        End With
        
         With .Cells(2, rng.Columns.Count)
           .Value = author & "/" & Format(Now, "dd.mm.yyyy")
           .HorizontalAlignment = xlRight
        End With
        
    End With
    
    With Application
        .ScreenUpdating = True
    End With
    
    Set addCaption = rng
    
End Function

Function formatWithCaption(Optional ByRef rng As Range, Optional hasHeading As Boolean = True, Optional hasFooting As Boolean = False, Optional keepOldFormat As Boolean = False, Optional bgDark As Long, Optional bgLight As Long, Optional ByVal category As String = "Arbeitspapier", Optional ByVal theme As String = "Demo", Optional ByVal index As String = "M3-2018/001", Optional ByVal author As String = "Qiou Yang") As Range
    addCaption rng, category, theme, index, author, bgDark
    formatRng rng, hasHeading, hasFooting, keepOldFormat, bgDark, bgLight
    
    Set formatWithCaption = rng
End Function


' one row or one column
' mergeCells with the same content
Function mergeCells(Optional ByRef rng As Range, Optional ByVal orient As String = "v")
    
     With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
     If IsMissing(rng) Or TypeName(rng) = "Nothing" Then
        Set rng = Application.Intersect(ActiveSheet.UsedRange, Selection)
    End If
    
    Dim i, tmpVal
    Dim thisC As Range
    Dim nextC As Range
    Dim prevC As Range
    Dim start As Range
    Dim ende As Range
    
    
    If rng.Cells.Count > 1 Then

        For i = rng.Cells.Count To 1 Step -1
        
            If orient = "v" Then
            
                Set thisC = rng.Cells(i, 1)
                Set nextC = rng.Cells(i - 1, 1)
                
                If i < rng.Cells.Count Then
                   Set prevC = rng.Cells(i + 1, 1)
                End If
             
            Else
                
                Set thisC = rng.Cells(1, i)
                Set nextC = rng.Cells(1, i - 1)
             
                If i < rng.Cells.Count Then
                   Set prevC = rng.Cells(1, i + 1)
                End If
            End If
        
            If i = rng.Cells.Count Then
                Set start = thisC
            ElseIf thisC.Value <> prevC.Value Then
                Set start = thisC
            End If
                
                
            If thisC.Value = nextC.Value Then
                If i = 1 Then
                    Set ende = thisC
                    tmpVal = thisC.Value
                    
                    With Range(start, ende)
                        .Merge
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With

                End If
            Else
                Set ende = thisC
                 With Range(start, ende)
                        .Merge
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                End With
            End If
        Next i
    End If
        
     With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    
    Set thisC = Nothing
    Set nextC = Nothing
    Set prevC = Nothing
    Set start = Nothing
    Set ende = Nothing

End Function

Function groupAndSum(ByVal targKeyCol1 As Integer, ByVal targKeyCol2 As Integer, Optional ByVal targValCol, Optional ByVal targRowBegine, Optional ByVal targRowEnd, Optional ByRef sht As Worksheet, Optional ByVal sorted As Boolean = False)
    
    If IsMissing(sht) Or TypeName(sht) = "Nothing" Then
        Set sht = ActiveSheet
    End If
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    
    With sht
        If IsMissing(targRowBegine) Then
            targRowBegine = 1
        End If
        
        If IsMissing(targRowEnd) Then
            targRowEnd = .Cells(.Rows.Count, targKeyCol2).End(xlUp).Row
        End If
        
        If IsMissing(targValCol) Then
            targValCol = targKeyCol2 + 1
        End If
        
        .Range(.Cells(targRowBegine, targKeyCol1), .Cells(targRowEnd, targKeyCol2)).ClearOutline
        
        Dim tmpPreviousRow As Integer
        Dim tmpCurrentRow As Integer
        
        tmpPreviousRow = targRowEnd
        tmpCurrentRow = tmpPreviousRow
        
         Do While tmpCurrentRow > targRowBegine
    
            tmpCurrentRow = .Cells(tmpCurrentRow, targKeyCol1).End(xlUp).Row
            
            If sorted Then
                With .Range(.Cells(tmpCurrentRow + 1, targKeyCol2), .Cells(tmpPreviousRow, targKeyCol2))
                    If .Cells.Count > 1 Then
                        .Sort Key1:=.Cells(1)
                    End If
                    .Rows.Group
                End With
            Else
                .Range(.Cells(tmpCurrentRow + 1, 1), .Cells(tmpPreviousRow, 1)).Rows.Group
            End If
            
            ' targValCol = 0  ignore sum
            If targValCol <> 0 Then
                .Cells(tmpCurrentRow, targValCol).Formula = "=SUM(" & .Cells(tmpCurrentRow + 1, targValCol).Address(0, 0) & ":" & .Cells(tmpPreviousRow, targValCol).Address(0, 0) & ")"
            End If
            
            tmpPreviousRow = tmpCurrentRow - 1
        Loop
    End With
    
     
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Function

Function addTextBoxComment(Optional content As String, Optional ByRef sht As Worksheet)
    
    If IsMissing(sht) Or TypeName(sht) = "Nothing" Then
        Set sht = ActiveSheet
    End If
    
    If IsMissing(content) Or content = "" Then
        content = "Prüfungshandlung : " & Chr(9) & "Durchlesen / Neuberechnung" & Chr(13) & "Prüfungsfeststellung : " & Chr(9) & "Keine Feststellung." & String(3, Chr(13)) & "Qiou Yang / " & Format(Now, "dd.mm.yyyy")
    End If
    
    With sht
        With .Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 480, 180)
            With .TextFrame2.TextRange.Characters
                .Text = content
                With .Font
                    .Size = pFontSizePlusPlus
                    .Name = pFont
                End With
            End With
            
            .Line.ForeColor.ObjectThemeColor = msoThemeColorAccent5
            .ShapeStyle = msoShapeStylePreset6
            
        End With
    End With

End Function

Function addBookmarks(Optional ByRef sht As Worksheet, Optional filterInclude, Optional filterExclude)
    
    If IsMissing(sht) Or TypeName(sht) = "Nothing" Then
        Set sht = ActiveSheet
    End If
    
    Dim w As Long
    Dim m As Long
    Dim h As Long
    Dim fs As Long
    
    Dim self As Object
    
    w = 180
    m = -15
    h = 50
    fs = 16
    
    Dim cnt As Long
    cnt = sht.Parent.Worksheets.Count - 1
    
    Dim arr()
    ReDim arr(0 To cnt)
    
    Dim i
    
    Dim shtNameArr()
    ReDim shtNameArr(0 To cnt)
    
    For i = 0 To cnt
        shtNameArr(i) = sht.Parent.Worksheets(i + 1).Name
    Next i
    
    
    If Not IsMissing(filterInclude) Then
        shtNameArr = filterArrayWith(shtNameArr, filterInclude, False)
    End If
    
    If Not IsMissing(filterExclude) Then
        shtNameArr = filterArrayWith(shtNameArr, filterExclude, True)
    End If
    
    Dim cnt1 As Long
    cnt1 = 0
    
    For i = 1 To cnt + 1
        If inArray(sht.Parent.Worksheets(cnt + 1 - i + 1).Name, shtNameArr) Then
            Set self = sht.Shapes.AddShape(msoShapeRound2SameRectangle, cnt1 * (w + m), 0, w, h)
             With self
                .ShapeStyle = msoShapeStylePreset22
                .TextFrame2.VerticalAnchor = msoAnchorMiddle
                .Fill.ForeColor.RGB = pBgLight
                .Fill.Solid
                
                 With .TextFrame2.TextRange.Characters
                    .ParagraphFormat.Alignment = msoAlignCenter
                    .Text = sht.Parent.Worksheets(cnt + 1 - i + 1).Name
                    
                    With .Font
                        .Fill.Solid
                        .Size = fs
                        .Name = pFont
                    End With
                 End With
                 
                 sht.Hyperlinks.Add Anchor:=self, Address:="", SubAddress:="'" & sht.Parent.Worksheets(cnt + 1 - i + 1).Name & "'" & "!B1"
                 
                 arr(i - 1) = .Name
            End With
            
            cnt1 = cnt1 + 1
        End If
    Next i
    
    Dim j, k
    
    With sht.Shapes.Range(arr).Group
        .Rotation = 270
        .Top = 0
        .Left = 0
        
        For Each i In sht.Parent.Worksheets
            If inArray(i.Name, shtNameArr) Then
                .Copy
                i.Paste i.Cells(1, 1)
            End If
        Next i
        
        .Delete
    End With
    
    
    
     For Each i In sht.Parent.Worksheets
        If inArray(i.Name, shtNameArr) Then
            For Each j In i.Shapes
                If j.Type = msoGroup Then
                    For Each k In j.GroupItems
                        ' if there are multiple group objects might be error.
                        If k.TextFrame2.TextRange.Characters.Text = i.Name Then
        
                            With k.Fill
                                .ForeColor.RGB = pBgDark
                                .Solid
                            End With
                            
                            
                             With k.TextFrame2.TextRange.Characters.Font.Fill
                                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                            End With
    
                        End If
                    Next k
                End If
            Next j
        End If
    Next i
End Function

Public Function removeGroupShapes(Optional ByRef wb As Workbook)

    Dim f1, s
    
    If IsMissing(wb) Or TypeName(wb) = "Nothing" Then
        Set wb = ActiveSheet.Parent
    End If
    
    
    For Each f1 In wb.Worksheets
        For Each s In f1.Shapes
            If s.Type = msoGroup Then
                s.Delete
            End If
        Next s
    Next f1

End Function

Private Function inArray(f, ByRef arr) As Boolean
    Dim res As Boolean
    Dim e
    
    res = False
    
    For Each e In arr
        If e = f Then
            res = True
            Exit For
        End If
    Next e
    
    inArray = res
End Function

Function filterArrayWith(ByRef arr, Optional f, Optional ByVal exclude As Boolean = False) As Variant
        Dim res()
        ReDim res(LBound(arr) To UBound(arr))
        
        Dim cnt As Long
        cnt = LBound(arr)
        
        Dim i
        
        If TypeName(f) = "IRegExp2" Then
            For Each i In arr
                If f.test(i) = Not exclude Then
                    res(cnt) = i
                    cnt = cnt + 1
                End If
            Next i
        ElseIf IsArray(f) Then
            For Each i In arr
                If inArray(i, f) = Not exclude Then
                    res(cnt) = i
                    cnt = cnt + 1
                End If
            Next i
        Else
            Err.Raise 9987, , "Unkown ParameterType filterInclude: Should be either RegExp or Array"
        End If
        
        If cnt > LBound(arr) Then
            ReDim Preserve res(LBound(arr) To cnt - 1)
            filterArrayWith = res
        Else
            filterArrayWith = Array()
        End If
    
End Function



Function dataFormat(Optional ByRef rng As Range, Optional ByVal fmtStr As String = "#,##0.00", Optional ByVal multiLines As Boolean = False) As Range
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    If IsMissing(rng) Or TypeName(rng) = "Nothing" Then
        Set rng = Application.Intersect(ActiveSheet.UsedRange, Selection)
    End If
    
    rng.NumberFormat = fmtStr
    
    If multiLines Then
        rng.WrapText = True
    End If
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    
    Set dataFormat = rng
    
End Function

Function formatAsCurrency(Optional ByRef rng As Range, Optional ByRef currencySymbol As String = "") As Range
    Dim res As Range
    
    Set res = dataFormat(rng, currencySymbol & " #,##0.00")
    res.HorizontalAlignment = xlRight
    
    Set formatAsCurrency = res
    Set res = Nothing
End Function

Function formatAsCurrencyThousand(Optional ByRef rng As Range, Optional ByRef currencySymbol As String = "", Optional ByVal decimalPos As Integer = 1) As Range
    Dim res As Range
    
    Set res = dataFormat(rng, currencySymbol & " #,." & String(decimalPos, "0"))
    res.HorizontalAlignment = xlRight
    
    Set formatAsCurrencyThousand = res
    Set res = Nothing
End Function

' overwrite the content if specify the year value
Function formatAsMonatAndYear(Optional ByRef rng As Range, Optional ByVal year As Integer, Optional ByVal startMonth As Integer, Optional ByVal rowHeight As Integer) As Range
    Dim res As Range
    Set res = dataFormat(rng, "mmm" & Chr(10) & "yyyy", True)
    
    Dim e
    Dim cnt As Integer
    
    If rowHeight = 0 Or IsMissing(rowHeight) Then
        rowHeight = pRowHeightCaption
    End If
    
    rng.rowHeight = rowHeight
    rng.HorizontalAlignment = xlCenter
    
    If Not (IsMissing(year) Or year = 0) Then
        
        If startMonth = 0 Or IsMissing(startMonth) Then
            startMonth = 1
        End If
    
        For Each e In rng.Cells
            e.Value = DateSerial(year, startMonth + cnt, 1)
            cnt = cnt + 1
        Next e
    End If

    Set formatAsMonatAndYear = res
    Set res = Nothing
End Function
