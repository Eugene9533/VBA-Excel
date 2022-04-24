Attribute VB_Name = "DateSubFunction"
Public Mon As Integer, StartMonth As Integer
Public CurrentYear As String, StartYear As String
Public CurrentDay As Integer, StartDay As Integer
Public ActiveButton As String
Public needDate As Date
Public leftOffset As Integer
Public topOffset As Integer
Sub SetNeedDateGlobalPerem(dateParam As Date)
    needDate = dateParam
End Sub
Function StringMonth(Month As Integer)
         StringMonth = Trim(Mid("январь  ‘евраль ћарт    јпрель  ћай     »юнь    »юль    јвгуст  —ент€брьќкт€брь Ќо€брь  ƒекабрь ", (Month - 1) * 8 + 1, 8))
End Function
Function Visok(Year As Long) As Boolean
   If Year Mod 400 = 0 Then Visok = True: Exit Function
   If Year Mod 100 = 0 Then Visok = False: Exit Function
   If Year Mod 4 = 0 Then
      Visok = True
   Else
      Visok = False
   End If
End Function
Function DayOfWeek(Day As Integer, Month As Integer, Year As Long)
  Dim n As Integer
  If Month < 3 Then
     If Visok(Year) Then
        n = 1
     Else
        n = 2
     End If
  Else
    n = 0
  End If
  DayOfWeek = (Fix(365.25 * Year) + Fix(30.56 * Month) + Day + n) Mod 7
End Function

'Function NameDayOfWeek(DayOfWeek)
'  NameDayOfWeek = Trim(Mid("ѕн¬т—р„тѕт—б¬с", (DayOfWeek) * 2 + 1, 2))
'End Function
Function NameDayOfWeek(DayOfWeek)
  NameDayOfWeek = Trim(Mid("ѕонедельник¬торник    —реда      „етверг    ѕ€тница    —уббота    ¬оскресенье", (DayOfWeek) * 11 + 1, 11))
End Function
Function MonthForDay(Month As Integer)
    MonthForDay = Trim(Mid("€нвар€  феврал€ марта   апрел€  ма€     июн€    июл€    августа сент€бр€окт€бр€ но€бр€  декабр€ ", (Month - 1) * 8 + 1, 8))
End Function
