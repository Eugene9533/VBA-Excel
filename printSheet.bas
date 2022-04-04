Attribute VB_Name = "printSheet"
Private Sub Workbook_Open()
Sub printSheet()
    With Sheets("Application").PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Sheets("Application").PageSetup.PrintArea = "$A$1:$BI$44"
    With Sheets("Application").PageSetup
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .CenterVertically = False
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Zoom = 63
    End With
    Sheets("Application").PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub
