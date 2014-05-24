
<!-- saved from url=(0055)http://profs.info.uaic.ro/~rvlad/lab/msoffice/lista.bas -->
<html><head><meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1"><style type="text/css"></style></head><body><pre style="word-wrap: break-word; white-space: pre-wrap;">Attribute VB_Name = "Module1"
Sub lista()
    Dim wa As Word.Application
    Dim wd As Word.Document
    Set wa = New Word.Application
    wa.Visible = True
    Set wd = wa.Documents.Add
    With wd.PageSetup
        .TopMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
    End With
    wa.Selection.TypeText ("Cãrþi")
    wa.Selection.TypeParagraph
    wd.Paragraphs(1).Range.Font.Bold = True
End Sub
</pre></body></html>