 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class FormatUtil
'                                          to implement the format uniformly
'@author                                   Qiou Yang
'@lastUpdate                               21.09.2018
'                                          add default theme and several number formats
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
        Set rng = Application.intersect(ActiveSheet.UsedRange, Selection)
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
        contentEndRow = .Rows.count
        
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
            With .Rows(.Rows.count)
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
            With .Rows(.Rows.count)
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
        Set rng = Application.intersect(ActiveSheet.UsedRange, Selection)
    End If
    
     If IsMissing(bgColor) Or bgColor = 0 Then
        bgColor = pBgDark
    End If
    
    Dim fontColorWhite As Long
    fontColorWhite = 16777215

    rng.Worksheet.Cells(1, rng.Cells(1, 1).Column).Resize(pBlankRowsUnterCaption + 2, rng.Columns.count).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    With rng.Worksheet.Cells(1, rng.Cells(1, 1).Column).Resize(2, rng.Columns.count)
        .Interior.Color = bgColor
        .Font.Name = pFont
        .Font.Size = pFontSizePlus
        .Font.Color = fontColorWhite
        
        .rowHeight = pRowHeightFooting
        
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        
        
        With .Cells(1, 1)
            .value = category
        End With
        
        With .Cells(2, 1)
           .value = theme
           .Font.Bold = True
           .Font.Size = pFontSizePlusPlus
        End With
        
        With .Cells(1, rng.Columns.count)
           .value = index
           .HorizontalAlignment = xlRight
        End With
        
         With .Cells(2, rng.Columns.count)
           .value = author & "/" & format(Now, "dd.mm.yyyy")
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
        Set rng = Application.intersect(ActiveSheet.UsedRange, Selection)
    End If
    
    Dim i, tmpVal
    Dim thisC As Range
    Dim nextC As Range
    Dim prevC As Range
    Dim start As Range
    Dim ende As Range
    
    
    If rng.Cells.count > 1 Then

        For i = rng.Cells.count To 1 Step -1
        
            If orient = "v" Then
            
                Set thisC = rng.Cells(i, 1)
                Set nextC = rng.Cells(i - 1, 1)
                
                If i < rng.Cells.count Then
                   Set prevC = rng.Cells(i + 1, 1)
                End If
             
            Else
                
                Set thisC = rng.Cells(1, i)
                Set nextC = rng.Cells(1, i - 1)
             
                If i < rng.Cells.count Then
                   Set prevC = rng.Cells(1, i + 1)
                End If
            End If
        
            If i = rng.Cells.count Then
                Set start = thisC
            ElseIf thisC.value <> prevC.value Then
                Set start = thisC
            End If
                
                
            If thisC.value = nextC.value Then
                If i = 1 Then
                    Set ende = thisC
                    tmpVal = thisC.value
                    
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

Function groupAndSum(ByVal targKeyCol1 As Integer, ByVal targKeyCol2 As Integer, Optional ByVal targValCol, Optional ByVal targRowBegine, Optional ByVal targRowEnd, Optional ByRef Sht As Worksheet, Optional ByVal sorted As Boolean = False)
    
    If IsMissing(Sht) Or TypeName(Sht) = "Nothing" Then
        Set Sht = ActiveSheet
    End If
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    
    With Sht
        If IsMissing(targRowBegine) Then
            targRowBegine = 1
        End If
        
        If IsMissing(targRowEnd) Then
            targRowEnd = .Cells(.Rows.count, targKeyCol2).End(xlUp).row
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
    
            tmpCurrentRow = .Cells(tmpCurrentRow, targKeyCol1).End(xlUp).row
            
            If sorted Then
                With .Range(.Cells(tmpCurrentRow + 1, targKeyCol2), .Cells(tmpPreviousRow, targKeyCol2))
                    If .Cells.count > 1 Then
                        .sort Key1:=.Cells(1)
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

Function addTextBoxComment(Optional content As String, Optional ByRef Sht As Worksheet)
    
    If IsMissing(Sht) Or TypeName(Sht) = "Nothing" Then
        Set Sht = ActiveSheet
    End If
    
    If IsMissing(content) Or content = "" Then
        content = "Prüfungshandlung : " & Chr(9) & "Durchlesen / Neuberechnung" & Chr(13) & "Prüfungsfeststellung : " & Chr(9) & "Keine Feststellung." & String(3, Chr(13)) & "Qiou Yang / " & format(Now, "dd.mm.yyyy")
    End If
    
    With Sht
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

Function dataFormat(Optional ByRef rng As Range, Optional ByVal fmtStr As String = "#,##0.00", Optional ByVal multiLines As Boolean = False) As Range
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    If IsMissing(rng) Or TypeName(rng) = "Nothing" Then
        Set rng = Application.intersect(ActiveSheet.UsedRange, Selection)
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
            e.value = DateSerial(year, startMonth + cnt, 1)
            cnt = cnt + 1
        Next e
    End If

    Set formatAsMonatAndYear = res
    Set res = Nothing
End Function
