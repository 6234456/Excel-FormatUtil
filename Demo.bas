Sub trail()

    Dim f As New FormatUtil
    
    f.setTheme FormatTheme.fmtOrange
    f.formatAsMonatAndYear rng:=f.formatWithCaption(rng:=Selection, hasFooting:=True).Rows(1), year:=2019

End Sub
