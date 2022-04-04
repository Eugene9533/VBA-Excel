Sub sendMail()
    Dim x As String
    Dim oOutlApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sTblBody As String, sAttachment As String
    Dim rDataR As Range
    Dim IsOultOpen As Boolean
    
    ActiveWorkbook.Save
    Application.DisplayAlerts = False
    strPath = ActiveWorkbook.Path & "\Temp\"
    On Error Resume Next
    x = GetAttr(strPath) And 0
    If Err = 0 Then
    strdate = Format(Now, "yyyy/mm")
    ActiveWorkbook.SaveAs Filename:=strPath & strdate & ".xlsm", FileFormat:= _
    xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Application.DisplayAlerts = True
    End If
    
    Application.ScreenUpdating = False
    On Error Resume Next
    Set oOutlApp = GetObject(, "Outlook.Application")
    If Err = 0 Then
        IsOultOpen = True
    Else
        Err.Clear
        Set oOutlApp = CreateObject("Outlook.Application")
    End If
    oOutlApp.Session.Logon
    Set objMail = oOutlApp.CreateItem(0)
    If Err.Number <> 0 Then Set oOutlApp = Nothing: Set objMail = Nothing: Exit Sub
       
    With ActiveWorkbook.Sheets("Form")
        sTo = .Range("BB1").Value
        sSubject = "Çàÿâêà íà " + "'" + .Range("E4").Value + "'"
        sBody = .Range("BA3").Value
        sBody = Replace(sBody, Chr(10), "<br />")
        sBody = Replace(sBody, vbNewLine, "<br />")
        sBody = "<span style=""font-size: 14px; font-family: Arial"">" & sBody & "</span>"
        Set rDataR = Sheets("Application").Range("$A$1:$BI$44")
        sTblBody = ConvertRngToHTM(rDataR)
        sBody = Replace(sBody, "{TABLE}", sTblBody)
    End With
    
    With objMail
        .To = sTo 'àäðåñ ïîëó÷àòåëÿ
        .Subject = sSubject 'òåìà ñîîáùåíèÿ
        .BodyFormat = 2
        .HTMLBody = sBody
        .Attachments.Add ActiveWorkbook.FullName
        .display
    End With
    
    If IsOultOpen = False Then oOutlApp.Quit
    Set oOutlApp = Nothing: Set objMail = Nothing
    DoEvents
End Sub

Function ConvertRngToHTM(rng As Range)
    Dim fso As Object, ts As Object
    Dim sF As String, resHTM As String
    Dim wbTmp As Workbook
 
    sF = Environ("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    rng.Copy
    Set wbTmp = Workbooks.Add(1)
    With wbTmp.Sheets(1)
        .Cells(1).PasteSpecial xlPasteColumnWidths
        .Cells(1).PasteSpecial xlPasteValues
        .Cells(1).PasteSpecial xlPasteFormats
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
    With wbTmp.PublishObjects.Add( _
         SourceType:=xlSourceRange, Filename:=sF, _
         Sheet:=wbTmp.Sheets(1).Name, Source:=wbTmp.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sF).OpenAsTextStream(1, -2)
    resHTM = ts.ReadAll
    ts.Close
    ConvertRngToHTM = Replace(resHTM, "align=center x:publishsource=", "align=left x:publishsource=")
    wbTmp.Close False
    Kill sF
    Set ts = Nothing: Set fso = Nothing
    Set wbTmp = Nothing
End Function

Function RangeToTextTable(rng As Range)
    Dim lr As Long, lc As Long, arr
    Dim res As String, rh()
    Dim lSpaces As Long, s As String
     
    arr = rng.Value
    If Not IsArray(arr) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = rng.Value
    End If
    ReDim rh(1 To UBound(arr, 2))
    For lr = 1 To UBound(arr, 1)
        For lc = 1 To UBound(arr, 2)
            If Len(arr(lr, lc)) > rh(lc) Then
                rh(lc) = Len(arr(lr, lc))
            End If
        Next
    Next
    For lr = 1 To UBound(arr, 1)
        For lc = 1 To UBound(arr, 2)
            s = arr(lr, lc)
            lSpaces = rh(lc) - Len(s)
            If lSpaces > 0 Then
                s = s & Space(lSpaces)
            End If
            If lc = 1 Then
                res = res & s
            Else
                res = res & vbTab & s
            End If
        Next
        res = res & vbNewLine
    Next
    RangeToTextTable = res
End Function
