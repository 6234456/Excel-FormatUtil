Sub trail()

    Dim f As New FormatUtil
    
    f.setTheme FormatTheme.fmtOrange
    f.formatAsMonatAndYear rng:=f.formatWithCaption(rng:=Selection, hasFooting:=True).Rows(1), year:=2019

End Sub

Private Sub trail1()

    Dim f As New FormatUtil

    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    reg.Pattern = "^Blatt"
    
    Dim e
    Dim k
    
    For Each e In Worksheets
        For Each k In e.Shapes
            k.Delete
        Next k
    Next e
    
    f.addBookmarks , reg

End Sub


Private Sub demo()

    Dim f As New FormatUtil
    
    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    reg.Pattern = "^1yang"
    
   For Each i In f.filterArrayWith(Array("qiou", "yang", "yang2"), reg, True)
    
    Debug.Print i
    
   Next i
End Sub


Sub demo1_fmtReportingPackage()

    Dim cnt As Integer
    cnt = 0
    
    Dim a, i
    
    a = Array("Übersicht", "Bilanz", "GuV", "Cashflow")

    Dim f As New FormatUtil
    
    f.setTheme FormatTheme.fmtBlue
    
    For Each i In a
         Worksheets.Add after:=Worksheets(Worksheets.Count)
         
         With Worksheets(Worksheets.Count)
            .Name = i
            With .Cells(1, 2)
                .Resize(10, 10).Value = 1
                f.formatWithCaption rng:=.Resize(10, 10), hasFooting:=True, theme:=i, category:="Demo", index:="M3-2018/" & Format(cnt, "000")
            End With
         End With
         
         cnt = cnt + 1
    Next i
    
    f.addBookmarks filterInclude:=a

End Sub

