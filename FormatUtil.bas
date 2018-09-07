 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class FormatUtil
'                                          to implement the format uniformly
'@author                                   Qiou Yang
'@lastUpdate                               07.09.2018
'@TODO                                   ´
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

Private Sub Class_Initialize()
    pFont = "BahnSchrift"
    pFontSize = 10
    pFontSizePlus = pFontSize + 1
    pFontSizePlusPlus = pFontSize + 4
    pRowHeightCaption = 48
    pRowHeightFooting = 35
    pRowHeight = 24
    pBlankRowsUnterCaption = 2
End Sub

Function formatRng(Optional ByRef rng As Range, Optional hasHeading As Boolean = True, Optional hasFooting As Boolean = False, Optional bgDark As Long = 11892015, Optional bgLight As Long = 16247773)
    
    With Application
        .ScreenUpdating = False
    End With
    
    If IsMissing(rng) Or TypeName(rng) = "Nothing" Then
        Set rng = Application.Intersect(ActiveSheet.UsedRange, Selection)
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
    
        .ClearFormats
        ActiveWindow.DisplayGridlines = False
        
        ' Alignment
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

        ' Font
        .Font.Name = pFont
        .Font.Size = pFontSize
        .Font.Color = fontColorBlack
        
        ' Row Height
        .RowHeight = pRowHeight
        
        If hasHeading Then
            With .Rows(1)
                With .Font
                    .Size = pFontSizePlus
                    .Color = fontColorWhite
                End With
                .Interior.Color = bgDark
                
                .RowHeight = pRowHeightCaption
                
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
                
                .RowHeight = pRowHeightFooting
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
                    .ThemeColor = 1
                End With
            End With
        End If
        
        .Borders(xlInsideVertical).Weight = xlThin

        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
    End With
    
    
     With Application
        .ScreenUpdating = True
    End With
    
End Function


' caption occupies the first two rows above the table
' @param category : up-left   can be for example  Recherche / Arbeitspapier / Neuberechnung / Interview
Function addCaption(Optional ByRef rng As Range, Optional ByVal category As String = "Arbeitspapier", Optional ByVal theme As String = "Demo", Optional ByVal index As String = "M3-2018/001", Optional ByVal author As String = "Qiou Yang", Optional bgColor As Long = 11892015)
    
    With Application
        .ScreenUpdating = False
    End With
    
     If IsMissing(rng) Or TypeName(rng) = "Nothing" Then
        Set rng = Application.Intersect(ActiveSheet.UsedRange, Selection)
    End If
    
    Dim fontColorWhite As Long
    fontColorWhite = 16777215

    rng.Worksheet.Cells(1, rng.Cells(1, 1).Column).Resize(pBlankRowsUnterCaption + 2, rng.Columns.Count).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    With rng.Worksheet.Cells(1, rng.Cells(1, 1).Column).Resize(2, rng.Columns.Count)
        .Interior.Color = bgColor
        .Font.Name = pFont
        .Font.Size = pFontSizePlus
        .Font.Color = fontColorWhite
        
        .RowHeight = pRowHeightFooting
        
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
    
End Function

Function formatWithCaption(Optional ByRef rng As Range, Optional hasHeading As Boolean = True, Optional hasFooting As Boolean = False, Optional bgDark As Long = 11892015, Optional bgLight As Long = 16247773, Optional ByVal category As String = "Arbeitspapier", Optional ByVal theme As String = "Demo", Optional ByVal index As String = "M3-2018/001", Optional ByVal author As String = "Qiou Yang")
    addCaption rng, category, theme, index, author, bgDark
    formatRng rng, hasHeading, hasFooting, bgDark, bgLight
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
            targRowEnd = .Cells(.Rows.Count, targKeyCol2).End(xlUp).row
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
