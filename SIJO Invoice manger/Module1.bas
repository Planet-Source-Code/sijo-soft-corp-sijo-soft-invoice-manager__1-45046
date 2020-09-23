Attribute VB_Name = "Module1"
Dim units(0 To 20) As String
Dim tens(1 To 10) As String
Dim ident(0 To 4) As String
Dim lon As Integer, tempword As String
Public Function inttoword(number As String) As String
    populatewords
    tempword = ""
    number = CDbl(number)
    lon = Len(number)
    If lon < 10 Then
        tempword = Crores(number)
        inttoword = tempword
    Else
        inttoword = "Numbers Greater Than Ten Crore Not Currently Supported!"
    End If
End Function
Private Sub populatewords()
    units(1) = "One"
    units(2) = "Two"
    units(3) = "Three"
    units(4) = "Four"
    units(5) = "Five"
    units(6) = "Six"
    units(7) = "Seven"
    units(8) = "Eight"
    units(9) = "Nine"
    units(10) = "Ten"
    units(11) = "Eleven"
    units(12) = "Tweleve"
    units(13) = "Thirteen"
    units(14) = "Fourteen"
    units(15) = "Fifteen"
    units(16) = "Sixteen"
    units(17) = "Seventeen"
    units(18) = "Eighteen"
    units(19) = "Nineteen"
    tens(1) = "Ten"
    tens(2) = "Twenty"
    tens(3) = "Thirty"
    tens(4) = "Fourty"
    tens(5) = "Fifty"
    tens(6) = "Sixty"
    tens(7) = "Seventy"
    tens(8) = "Eighty"
    tens(9) = "Ninety"
    ident(1) = "Hundred"
    ident(2) = "Thousand"
    ident(3) = "Lac"
    ident(4) = "Crore"
End Sub
Private Function tenfunc(number As String) As String
Dim tempint As Integer
Dim sp As String
tempint = 0: sp = ""
tempint = CInt(Right(number, 2))
    number = CStr(tempint)
    If tempint < 20 Then
        tenfunc = units(tempint)
    ElseIf tempint > 19 Then
        If CInt(Right(number, 1)) <> 0 Then sp = " " Else sp = ""
        tenfunc = tens(CInt(Left(number, 1))) & sp & units(CInt(Right(number, 1)))
    End If
End Function
Private Function hundreds(number As String) As String
Dim tempint As Integer
tempint = CInt(Right(number, 3))
number = CStr(tempint)
    If Len(number) = 3 Then
        hundreds = units(CInt(Left(number, 1))) & " " & ident(1) & " " & tenfunc(Right(number, 2))
    ElseIf Len(number) < 3 Then
        hundreds = tenfunc(number)
    End If
End Function
Private Function thousands(number As String) As String
Dim tempint As Double
Dim tempstr As String
tempstr = ""
tempint = CDbl(number)
number = tempint
    If Len(number) = 5 Then
        tempstr = Left(number, 2)
        thousands = tenfunc(tempstr)
        If CInt(tempstr) > 0 Then thousands = thousands & " " & ident(2)
        thousands = thousands & " " & hundreds(Right(number, 3))
    ElseIf Len(number) = 4 Then
        tempstr = Left(number, 1)
        thousands = tenfunc(tempstr)
        If CInt(tempstr) > 0 Then thousands = thousands & " " & ident(2)
        thousands = thousands & " " & hundreds(Right(number, 3))
    ElseIf Len(number) < 4 Then
        thousands = hundreds(number)
    End If
End Function
Private Function lacs(number As String) As String
Dim tempint As Double
Dim tempstr As String
tempstr = ""
tempint = CDbl(number)
number = tempint
    If Len(number) = 7 Then
        tempstr = Left(number, 2)
        lacs = tenfunc(tempstr)
        If CInt(tempstr) > 0 Then lacs = lacs & " " & ident(3)
        lacs = lacs & " " & thousands(Right(number, 5))
    ElseIf Len(number) = 6 Then
        tempstr = Left(number, 1)
        lacs = tenfunc(tempstr)
        If CInt(tempstr) > 0 Then lacs = lacs & " " & ident(3)
        lacs = lacs & " " & thousands(Right(number, 5))
    ElseIf Len(number) < 6 Then
        lacs = thousands(number)
    End If
End Function
Private Function Crores(number As String) As String
Dim tempint As Double
Dim tempstr As String
tempstr = ""
tempint = CDbl(number)
number = tempint
    If Len(number) = 9 Then
        tempstr = Left(number, 2)
        Crores = tenfunc(tempstr)
        If CInt(tempstr) > 0 Then Crores = Crores & " " & ident(4)
        Crores = Crores & " " & lacs(Right(number, 7))
    ElseIf Len(number) = 8 Then
        tempstr = Left(number, 1)
        Crores = tenfunc(tempstr)
        If CInt(tempstr) > 0 Then Crores = Crores & " " & ident(4)
        Crores = Crores & " " & lacs(Right(number, 7))
    ElseIf Len(number) < 8 Then
        Crores = lacs(number)
    End If
End Function


Public Sub main()
MsgBox "I want to convert it to MS Access Database, Please help me and Re-Submit as 'Updated SIJO Soft Invoice Manager"
FrmSplash.Show
End Sub
