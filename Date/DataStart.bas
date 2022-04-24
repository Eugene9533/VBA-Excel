Attribute VB_Name = "Module4"
Sub DataStart()
    Dim dateObjectForChange As Object
    Set dateObjectForChange = Cells(48, 5)
    If IsDate(dateObjectForChange.Value) Then Call DateSubFunction.SetNeedDateGlobalPerem(dateObjectForChange.Value) _
    Else Call DateSubFunction.SetNeedDateGlobalPerem(Date)
    DateForm.Show
    dateObjectForChange.Value = Format(DateSubFunction.needDate)
End Sub
