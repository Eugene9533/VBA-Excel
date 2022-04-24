VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateForm 
   Caption         =   "Календарь"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5550
   OleObjectBlob   =   "DateForm.frx":0000
End
Attribute VB_Name = "DateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartDayWeek As Integer, CountFebrary As Integer
Dim Silence As Boolean
Dim CountDays As Integer
Dim ctl As control
Dim Labels(0 To 41) As New DateClass
Dim lblYears(0 To 3) As New DateYearClass
Private Sub cmbMonth_Change()
   If Silence = True Then Exit Sub
   Silence = True
   Mon = cmbMonth.ListIndex + 1 '(InStr(1, "Январь  Февраль Март    Апрель  Май     Июнь    Июль    Август  СентябрьОктябрь Ноябрь  Декабрь ", cmbMonth.Value) - 1) / 8 + 1
   scbMonth.Value = Mon
   Call Refresh(CurrentDay, Mon, CLng(CurrentYear))
   Silence = False
   Me.Caption = NameDayOfWeek(DayOfWeek(CurrentDay, Mon, CLng(CurrentYear))) & " " & CurrentDay & " " & MonthForDay(Mon) & " " & CurrentYear
End Sub
Private Sub cmbMonth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
       If KeyCode = vbKeyV And Shift = 2 Then KeyCode = 0
       If KeyCode = vbKeyDelete Then KeyCode = 0
       If KeyCode = vbKeyBack Then KeyCode = 0
End Sub
Private Sub cmbMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub Label18_Click()

End Sub

Private Sub sbtSelectYear_Change()
    If Silence = True Then Exit Sub
    Silence = True
    On Error Resume Next
    CurrentYear = str(sbtSelectYear.Value)
    tbxYear.Text = CurrentYear
    lblYear1.Caption = Mid(CurrentYear, Len(CurrentYear), 1)
    lblYear2.Caption = Mid(CurrentYear, Len(CurrentYear) - 1, 1)
    lblYear3.Caption = Mid(CurrentYear, Len(CurrentYear) - 2, 1)
    lblYear4.Caption = Mid(CurrentYear, Len(CurrentYear) - 3, 1)
    Call Refresh(CurrentDay, Mon, CLng(CurrentYear))
    Silence = False
    Me.Caption = NameDayOfWeek(DayOfWeek(CurrentDay, Mon, CLng(CurrentYear))) & " " & CurrentDay & " " & MonthForDay(Mon) & " " & CurrentYear
End Sub
Private Sub sbtSelectYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Silence = True Then Exit Sub
       Silence = True
       Call Refresh(CurrentDay, Mon, CLng(CurrentYear))
       Silence = False
End Sub
Private Sub scbMonth_Change()
        If Silence = True Then Exit Sub
        Silence = True
        Mon = scbMonth.Value
        cmbMonth.ListIndex = scbMonth.Value - 1
        Call Refresh(CurrentDay, Mon, CLng(CurrentYear))
        Silence = False
        Me.Caption = NameDayOfWeek(DayOfWeek(CurrentDay, Mon, CLng(CurrentYear))) & " " & CurrentDay & " " & MonthForDay(Mon) & " " & CurrentYear
End Sub
Private Sub tbxYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Call ChangeYear
End Sub
Private Sub tbxYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    End Sub
Private Sub tbxYear_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        If KeyCode = vbKeyV And Shift = 2 Then KeyCode = 0
        If KeyCode = vbKeyReturn Then
        Call ChangeYear
        End If
    End Sub
Private Sub UserForm_Initialize()
      Silence = True
    With cmbMonth
       .List = Split("Январь,Февраль,Март,Апрель,Май,Июнь,Июль,Август,Сентябрь,Октябрь,Ноябрь,Декабрь", ",")
       .Value = StringMonth(Month(DateSubFunction.needDate))
  End With
  If IsDate(DateSubFunction.needDate) Then
     CurrentDay = Day(DateSubFunction.needDate)
     CurrentYear = Trim(str(Year(DateSubFunction.needDate)))
     Mon = Month(DateSubFunction.needDate)
  Else
     CurrentDay = Day(Date)
     CurrentYear = Trim(str(Year(Date)))
     Mon = Month(Date)
  End If
  StartDay = CurrentDay
  StartMonth = Mon
  StartYear = CurrentYear
  scbMonth.Value = Mon
  sbtSelectYear.Value = CLng(CurrentYear)
  lblYear1.Caption = Mid(CurrentYear, Len(CurrentYear), 1)
  lblYear2.Caption = Mid(CurrentYear, Len(CurrentYear) - 1, 1)
  lblYear3.Caption = Mid(CurrentYear, Len(CurrentYear) - 2, 1)
  lblYear4.Caption = Mid(CurrentYear, Len(CurrentYear) - 3, 1)
  Call Refresh(CurrentDay, Mon, CLng(CurrentYear))
  Silence = False
  With Me
      .Caption = NameDayOfWeek(DayOfWeek(CurrentDay, Mon, CLng(CurrentYear))) & " " & CurrentDay & " " & MonthForDay(Mon) & " " & CurrentYear
      If DateSubFunction.leftOffset = 0 Then DateSubFunction.leftOffset = 350
      If DateSubFunction.topOffset = 0 Then DateSubFunction.topOffset = 250
      .left = DateSubFunction.leftOffset
      .top = DateSubFunction.topOffset
      DateSubFunction.leftOffset = 0
      DateSubFunction.topOffset = 0
  End With
End Sub
Sub Refresh(Day As Integer, Month As Integer, Year As Long)
    Dim CountDaysOfLastMonth As Integer
    CurrentDay = Day
    CountFebrary = IIf(Visok(Year), 29, 28)
    CountDays = Choose(Month, 31, CountFebrary, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    StartDayWeek = DayOfWeek(1, Month, Year)
    If Month = 1 Then
       CountDaysOfLastMonth = 31
    Else
       CountDaysOfLastMonth = Choose(Month - 1, 31, CountFebrary, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    End If
    For Each ctl In DateForm.Controls
        With ctl
            If .Tag = "DateButton" Then
               If .TabIndex < (StartDayWeek) Then
                  .Caption = CountDaysOfLastMonth - (StartDayWeek - ctl.TabIndex - 1)
                  .ForeColor = RGB(175, 175, 175)
                  .BackColor = Me.BackColor
                  '.SpecialEffect = fmSpecialEffectRaised
                  .SpecialEffect = fmSpecialEffectFlat
               ElseIf .TabIndex > (StartDayWeek + CountDays - 1) Then
                  .Caption = .TabIndex - (StartDayWeek + CountDays) + 1
                  .ForeColor = RGB(175, 175, 175)
                  .BackColor = Me.BackColor
                  '.SpecialEffect = fmSpecialEffectRaised
                  .SpecialEffect = fmSpecialEffectFlat
               Else
                  .Caption = .TabIndex - StartDayWeek + 1
                  If (.TabIndex + 1) Mod 7 = 0 Then
                     .ForeColor = RGB(255, 0, 0)
                  Else
                        .ForeColor = RGB(0, 0, 0)
                  End If
                  If .Caption = StartDay And Mon = StartMonth And CurrentYear = StartYear Then
                     .BackColor = RGB(255, 255, 255)
                     .SpecialEffect = fmSpecialEffectSunken
                     .ForeColor = RGB(0, 0, 0)
                  Else
                     .BackColor = Me.BackColor
                     .SpecialEffect = fmSpecialEffectRaised
                  End If
               End If
               On Error Resume Next
               Set Labels(.TabIndex).DateButton = ctl
            End If
            If ctl.Tag = "YearButton" Then
               On Error Resume Next
               Set lblYears(.TabIndex - 51).YearButton = ctl
            End If
        End With
    Next
    'Me.Caption = NameDayOfWeek(DayOfWeek(CurrentDay, Mon, CLng(CurrentYear))) & " " & CurrentDay & " " & MonthForDay(Mon) & " " & CurrentYear
End Sub
Sub SelectTextBox()
    For Each ctl In DateForm.Controls
        If ctl.Tag = "YearButton" Then
           ctl.Visible = False
           On Error Resume Next
           Set lblYears(ctl.TabIndex - 51).YearButton = ctl
        End If
    Next
    With tbxYear
         .Visible = True
         .SetFocus
    End With
End Sub
Sub MoveCursor()
    Dim MoveDay As Integer
    For Each ctl In DateForm.Controls
        With ctl
            If .Tag = "DateButton" Then
               If .ForeColor <> RGB(175, 175, 175) Then
                  If .Caption = ActiveButton Then
                     .ForeColor = RGB(0, 0, 255)
                     .SpecialEffect = fmSpecialEffectSunken
                     MoveDay = .Caption
                  Else
                     If (.TabIndex + 1) Mod 7 = 0 Then
                        .ForeColor = RGB(255, 0, 0)
                     Else
                        .ForeColor = RGB(0, 0, 0)
                     End If
                      If Not (.Caption = StartDay And Mon = StartMonth And CurrentYear = StartYear) Then
                         .SpecialEffect = fmSpecialEffectRaised
                      End If
                  End If
               End If
             End If
         End With
     Next
     Me.Caption = NameDayOfWeek(DayOfWeek(MoveDay, Mon, CLng(CurrentYear))) & " " & MoveDay & " " & MonthForDay(Mon) & " " & CurrentYear
End Sub
Sub ChangeYear()
       If Silence = True Then Exit Sub
       Silence = True
              On Error Resume Next
              lblYear1.Caption = Mid(tbxYear.Text, Len(tbxYear.Text), 1): If Not Err = 0 Then GoTo PASS
              lblYear2.Caption = Mid(tbxYear.Text, Len(tbxYear.Text) - 1, 1): If Not Err = 0 Then lblYear2.Caption = ""
              lblYear3.Caption = Mid(tbxYear.Text, Len(tbxYear.Text) - 2, 1): If Not Err = 0 Then lblYear3.Caption = ""
              lblYear4.Caption = Mid(tbxYear.Text, Len(tbxYear.Text) - 3, 1): If Not Err = 0 Then lblYear4.Caption = ""
              CurrentYear = tbxYear.Text
              sbtSelectYear.Value = CLng(CurrentYear)
       Call Refresh(CurrentDay, Mon, CLng(CurrentYear))
PASS:
              lblYear1.Visible = True
              lblYear2.Visible = True
              lblYear3.Visible = True
              lblYear4.Visible = True
              tbxYear.Visible = False
       Silence = False
       Me.Caption = NameDayOfWeek(DayOfWeek(CurrentDay, Mon, CLng(CurrentYear))) & " " & CurrentDay & " " & MonthForDay(Mon) & " " & CurrentYear
End Sub
