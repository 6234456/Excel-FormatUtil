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
