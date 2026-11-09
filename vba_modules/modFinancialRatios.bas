Attribute VB_Name = "modFinancialRatios"
Option Explicit

Public Function RatioNames() As Variant
    RatioNames = Split("Company|Period|Currency|Gross margin|EBITDA margin|EBIT margin|Net margin|ROE|ROIC|Current ratio|Quick ratio|Debt / equity|Debt / EBITDA|Interest cover|Revenue YoY|EBITDA YoY|Net income YoY|EV / revenue|EV / EBITDA|P/E", "|")
End Function

Private Function Ref(ByVal col As String, ByVal r As Long) As String
    Ref = "Financial_Statements!" & col & r
End Function

' COUNT distinguishes missing inputs from real zeros. NM is a policy outcome.
Private Function Guard(ByVal inputs As String, ByVal n As Long, ByVal invalid As String, ByVal expr As String) As String
    Guard = "=IF(COUNT(" & inputs & ")<" & n & ",""N/A"",IF(" & invalid & ",""N/M""," & expr & "))"
End Function

Private Function Quot(ByVal top As String, ByVal bottom As String, Optional ByVal nonnegative As Boolean = False) As String
    Dim invalid As String: invalid = bottom & "<=0"
    If nonnegative Then invalid = "OR(" & invalid & "," & top & "<0)"
    Quot = Guard(top & "," & bottom, 2, invalid, top & "/" & bottom)
End Function

Public Sub CalculateFinancialRatios()
    Dim w As Worksheet, src As Worksheet, r As Long, p As Long, c As Long, t As Double, finish As Long
    Dim num, den, f(), refs(), h, k As String, debt As String, capital As String, tr As String, e As String, d As Date
    t = Timer: Set src = S("Financial_Statements"): Set w = S("Financial_Ratios")
    finish = LastR(src): w.UsedRange.ClearContents: Headers w, 1, RatioNames()
    If finish < 2 Then Exit Sub
    ReDim f(1 To finish - 1, 1 To 20)
    For r = 2 To finish
        For c = 1 To 3: f(r - 1, c) = "=" & Ref(CStr(Array("A", "C", "E")(c - 1)), r): Next
        num = Array("I", "M", "O", "S"): den = Ref("G", r)
        For c = 0 To 3: f(r - 1, c + 4) = Quot(Ref(CStr(num(c)), r), CStr(den)): Next
        p = 0
        If r > 2 Then
            d = CDate(src.Cells(r, 4).Value2)
            If src.Cells(r - 1, 1).Value2 = src.Cells(r, 1).Value2 And src.Cells(r - 1, 5).Value2 = src.Cells(r, 5).Value2 Then
                If DateAdd("yyyy", 1, CDate(src.Cells(r - 1, 4).Value2)) = d Or DateAdd("yyyy", -1, d) = CDate(src.Cells(r - 1, 4).Value2) Then p = r - 1
            End If
        End If
        f(r - 1, 8) = "=""N/A""": f(r - 1, 9) = "=""N/A"""
        For c = 15 To 17: f(r - 1, c) = "=""N/A""": Next
        If p > 0 Then
            e = Ref("AC", r): k = Ref("AC", p)
            f(r - 1, 8) = Guard(Ref("S", r) & "," & e & "," & k, 3, "MIN(" & e & "," & k & ")<=0", Ref("S", r) & "/AVERAGE(" & e & "," & k & ")")
            capital = "(" & Ref("AB", p) & "+" & k & "-" & Ref("V", p) & ")": tr = Ref("R", r) & "/" & Ref("Q", r)
            f(r - 1, 9) = Guard(Ref("O", r) & "," & Ref("Q", r) & "," & Ref("R", r) & "," & Ref("AB", p) & "," & k & "," & Ref("V", p), 6, "OR(" & capital & "<=0," & Ref("Q", r) & "<=0," & Ref("R", r) & "<0," & Ref("R", r) & ">" & Ref("Q", r) & ")", Ref("O", r) & "*(1-" & tr & ")/" & capital)
            num = Array("G", "M", "S")
            For c = 0 To 2
                e = Ref(CStr(num(c)), r): k = Ref(CStr(num(c)), p)
                f(r - 1, c + 15) = Guard(e & "," & k, 2, k & "<=0", e & "/" & k & "-1")
            Next
        End If
        f(r - 1, 10) = Quot(Ref("U", r), Ref("Z", r))
        e = Ref("U", r): k = Ref("X", r): debt = Ref("Z", r)
        f(r - 1, 11) = Guard(e & "," & k & "," & debt, 3, debt & "<=0", "(" & e & "-" & k & ")/" & debt)
        f(r - 1, 12) = Quot(Ref("AB", r), Ref("AC", r), True)
        f(r - 1, 13) = Quot(Ref("AB", r), Ref("M", r), True)
        f(r - 1, 14) = Quot(Ref("O", r), Ref("P", r))
        f(r - 1, 18) = Quot(Ref("AG", r), Ref("G", r), True)
        f(r - 1, 19) = Quot(Ref("AG", r), Ref("M", r), True)
        f(r - 1, 20) = Quot(Ref("AF", r), Ref("S", r), True)
    Next
    w.Range("A2").Resize(finish - 1, 20).Formula = f: w.Calculate
    Audit "Ratios", "OK", finish - 1, "17 metrics per accepted filing.", Took(t)
End Sub
