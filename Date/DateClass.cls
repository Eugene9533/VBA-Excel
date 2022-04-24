VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents DateButton As MSForms.Label
Attribute DateButton.VB_VarHelpID = -1

Private Sub DateButton_MouseDown(ByVal button As Integer, _
    ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
       Call DateForm.ChangeYear
       If DateButton.ForeColor <> RGB(175, 175, 175) Then
          Call DateForm.Refresh(CInt(DateButton.Caption), Mon, CLng(CurrentYear))
       End If
    DateSubFunction.needDate = DateSerial(CInt(CurrentYear), Mon, CurrentDay)
    Unload DateForm
End Sub

Private Sub DateButton_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If DateButton.ForeColor <> RGB(175, 175, 175) Then
        ActiveButton = DateButton.Caption
        Call DateForm.MoveCursor
    End If
End Sub
