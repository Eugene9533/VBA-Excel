Private Sub Workbook_Open()
Function ExtractNfT(sWord As String, Optional Metod As Integer)
    Dim sSymbol As String, sInsertWord As String
    Dim i As Integer
 
    If sWord = "" Then ExtractNfT = "no data!": Exit Function
    sInsertWord = ""
    sSymbol = ""
    For i = 1 To Len(sWord)
        sSymbol = Mid(sWord, i, 1)
        If Metod = 1 Then
            If Not LCase(sSymbol) Like "*[0-9]*" Then
                If (sSymbol = "," Or sSymbol = "." Or sSymbol = " ") And i > 1 Then
                    If Mid(sWord, i - 1, 1) Like "*[0-9]*" And Mid(sWord, i + 1, 1) Like "*[0-9]*" Then
                        sSymbol = ""
                    End If
                End If
                sInsertWord = sInsertWord & sSymbol
            End If
        Else
            If LCase(sSymbol) Like "*[0-9¹:]*" Then
                If LCase(sSymbol) Like "*[.,]*" And i > 1 Then
                    If Not Mid(sWord, i - 1, 1) Like "*[0-9]*" Or Not Mid(sWord, i + 1, 1) Like "*[0-9]*" Then
                        sSymbol = " "
                    End If
                End If
                sInsertWord = sInsertWord & sSymbol
            End If
        End If
    Next i
    ExtractNfT = sInsertWord
End Function
