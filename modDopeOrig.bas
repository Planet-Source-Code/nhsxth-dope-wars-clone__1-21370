Attribute VB_Name = "Module1"
'Source code for Dopewars in VB
'Written by Matt Fredrikson, 2001
Type Drug
    name As String
    Price As Currency
    Qty As Integer
End Type

Type Character
    chName As String * 10
    Nationality As String * 25
    ID As Integer
    X As Integer
    Y As Integer
    Health As Integer
    Attack As Integer
    Move As Integer
    Val As Integer
    Coat As Integer
    narco(20) As Drug
End Type

Type Tiles
    Occupied As Boolean
    nOcc As Integer
    ocID(1 To 100) As Integer
    narco(20) As Drug
End Type

Type Game
    Nationality As String * 25
    plName As String * 35
    Money As Currency
    Chars(5) As Character
    nChars As Integer
    Debt As Currency
    Savings As Currency
    narco(20) As Drug
    gDay As Integer
    gCity As String * 35
    gLoc As String * 35
End Type

Type CostOffset
    dName As String
    Price As Currency
End Type

Type HighScores
    pname(10) As String * 35
    score(10) As Currency
End Type

Type GameText
    Text As String
End Type

Public ch(0 To 100) As Character
Public map(0 To 100, 0 To 100) As Tiles
Public team As String
Public nChars As Integer
Public gNat As String
Public gpName As String
Public dOffset(20) As CostOffset
Public Text2 As GameText
Global gtemp1 As Variant, gtemp2 As Variant, gtemp3 As Variant
Global lIntrest As Currency, bIntrest As Currency
Global DayNum As Integer, TotalDays As Integer, doneOpen As Integer
Global Money As Currency, Debt As Currency, Savings As Currency
Global doName As String, pCity As String, CurLoc As String
Global Coat As Integer
Dim ga As Game
Function Process(com As String)
    On Error Resume Next
    Text2.Text = ""
    Dim temp1 As Integer, temp2 As Integer, cid As Integer
    Dim tmpchars(20) As Character, nt As Integer, mode As Integer
    Dim tmptext As String, tmpwar As String, tmpdrug(20) As Drug
    Dim daOff As Integer
    If com = "cc" Then
        nChars = 1
        ch(nChars).ID = nChars
        ch(nChars).chName = "Lord"
        ch(nChars).Nationality = "Mr. Nice Guy"
        ch(nChars).Health = 100
        ch(nChars).Coat = 0
        Open "Drugs.txt" For Input As #1
        j = 0
        Do Until EOF(1)
            j = j + 1
            mode = 1
            Line Input #1, X
            If j = 1 Then nt = Val(X)
            For i = 1 To Len(X)
                If mode = 1 And Mid(X, i, 1) <> "," Then
                    tmptext = tmptext & Mid(X, i, 1)
                ElseIf mode = 1 And Mid(X, i, 1) = "," Then
                    tmpdrug(j).name = tmptext
                    tmptext = ""
                    mode = 2
                ElseIf mode = 2 Then
                    tmptext = tmptext & Mid(X, i, 1)
                    tmpdrug(j).Price = Val(tmptext)
                End If
            Next
            tmptext = ""
        Loop
        Close #1
        For i = 1 To j
            ch(nChars).narco(i).name = tmpdrug(i).name
            ch(nChars).narco(i).Qty = 0
        Next
    ElseIf com = "mc" Then
        temp1 = InputBox("Enter Destination X:", "Dest. X")
        temp2 = InputBox("Enter Destination Y:", "Dest. Y")
        cid = InputBox("Enter Character ID:", "Character ID")
        Call MoveChar(temp1, temp2, cid)
    ElseIf com = "ti" Then
        temp1 = InputBox("Enter Tile X:", "Tile X")
        temp2 = InputBox("Enter Tile Y:", "Tile Y")
        Call DispTiInfo(temp1, temp2)
    ElseIf com = "ci" Then
        temp1 = InputBox("Enter Character ID:", "Character ID")
        DispChInfo (temp1)
    ElseIf com = "bc" Then
        strtmp = InputBox("Enter Character Type:", "Character Type")
        tmpx = InputBox("Enter X:", "X Coordinate")
        tmpy = InputBox("Enter Y:", "Y Coordinate")
        Open "Characters.txt" For Input As #1
        j = 0
        Do Until EOF(1)
            j = j + 1
            mode = 1
            Line Input #1, X
            If j = 1 Then nt = Val(X)
            For i = 1 To Len(X)
                If mode = 1 And Mid(X, i, 1) <> "," Then
                    tmptext = tmptext & Mid(X, i, 1)
                ElseIf mode = 1 And Mid(X, i, 1) = "," Then
                    mode = 2
                ElseIf mode = 2 And Mid(X, i, 1) <> "," Then
                    tmpval = tmpval & Mid(X, i, 1)
                ElseIf mode = 2 And Mid(X, i, 1) = "," Then
                    tmpchars(j).Val = Val(tmpval)
                    tmpval = ""
                    mode = 3
                ElseIf mode = 3 Then
                    tmpwar = tmpwar & Mid(X, i, 1)
                    tmpchars(j).Attack = Val(tmpwar)
                End If
            Next
            tmpchars(j).chName = tmptext
            tmptext = ""
            tmpwar = ""
        Loop
        Close #1
        Open "Drugs.txt" For Input As #1
        j = 0
        Do Until EOF(1)
            j = j + 1
            mode = 1
            Line Input #1, X
            If j = 1 Then nt = Val(X)
            For i = 1 To Len(X)
                If mode = 1 And Mid(X, i, 1) <> "," Then
                    tmptext = tmptext & Mid(X, i, 1)
                ElseIf mode = 1 And Mid(X, i, 1) = "," Then
                    tmpdrug(j).name = tmptext
                    tmptext = ""
                    mode = 2
                ElseIf mode = 2 Then
                    tmptext = tmptext & Mid(X, i, 1)
                    tmpdrug(j).Price = Val(tmptext)
                End If
            Next
            tmptext = ""
        Loop
        Close #1
        nChars = nChars + 1
        For i = 1 To 20
            ch(nChars).narco(i).name = tmpdrug(i).name
            ch(nChars).narco(i).Qty = 0
            ch(nChars).narco(i).Price = 0
        Next
        For i = 1 To nt + 1
            If strtmp = Trim(tmpchars(i).chName) Then
                tmpcost = tmpchars(i).Val
                tmpwar = tmpchars(i).Attack
            End If
        Next
        If Money - tmpcost < 0 Then
            Y = MsgBox("You do not have enough money", vbInformation, "No Money")
            Exit Function
        End If
        With ch(nChars)
            .chName = strtmp
            .Attack = tmpwar
            .Health = 100
            .Move = 3
            .ID = nChars
            .Nationality = gNat
            .X = tmpx
            .Y = tmpy
        End With
        map(ch(nChars).X, ch(nChars).Y).Occupied = True
        map(ch(nChars).X, ch(nChars).Y).nOcc = map(ch(nChars).X, ch(nChars).Y).nOcc + 1
        map(ch(nChars).X, ch(nChars).Y).ocID(nChars) = ch(nChars).ID
        Money = Money - tmpcost
        DispChInfo (nChars)
        Text2.Text = Text2.Text & "You have $" & Money & " left."
        Text2.Text = Text2.Text & vbCrLf
    ElseIf com = "sg" Then
        Call SaveGame("C:\Windows\Desktop\dopeciv.txt")
        Text2.Text = Text2.Text & "Game Saved"
        Text2.Text = Text2.Text & vbCrLf
    ElseIf com = "lg" Then
        Call LoadGame("C:\Windows\Desktop\dopeciv.txt")
        Text2.Text = Text2.Text & "Loaded"
        Text2.Text = Text2.Text & vbCrLf
    ElseIf com = "sd" Then
        DrugN = gtemp1
        Quan = gtemp2
        cid = gtemp3
        If Quan = 0 Then
            Y = MsgBox("Impossible to sell nothing", vbInformation, "Error")
            Exit Function
        End If
        Text2.Text = Text2.Text & vbCrLf
        For i = 1 To 21
            If i = 21 Then
                Y = MsgBox("Drug does not exist", vbInformation, "Error")
                Exit Function
            End If
            If Trim(ch(cid).narco(i).name) = Trim(DrugN) Then Exit For
        Next
        For i = 1 To 20
            If ch(cid).narco(i).name = DrugN Then
                If ch(cid).narco(i).Qty < Quan Then
                    Y = MsgBox("You don't have that much " & DrugN & "!", vbInformation, "Error")
                    Exit Function
                End If
            End If
        Next
        Open "Drugs.txt" For Input As #1
        j = 0
        Do Until EOF(1)
            j = j + 1
            mode = 1
            Line Input #1, X
            If j = 1 Then nt = Val(X)
            For i = 1 To Len(X)
                If mode = 1 And Mid(X, i, 1) <> "," Then
                    tmptext = tmptext & Mid(X, i, 1)
                ElseIf mode = 1 And Mid(X, i, 1) = "," Then
                    tmpdrug(j).name = tmptext
                    tmptext = ""
                    mode = 2
                ElseIf mode = 2 Then
                    tmptext = tmptext & Mid(X, i, 1)
                    tmpdrug(j).Price = Val(tmptext)
                End If
            Next
            tmptext = ""
        Loop
        Close #1
        nt = nt + 1
        For i = 1 To nt
            If DrugN = Trim(tmpdrug(i).name) Then
                tmpcost = tmpdrug(i).Price
                tmptext = Trim(tmpdrug(i).name)
            End If
        Next
        For i = 1 To nt
            If dOffset(i).dName = DrugN Then
                tmpcost = tmpcost + dOffset(i).Price
            End If
        Next
        For i = 1 To nt
            If ch(cid).narco(i).name = DrugN Then
                ch(cid).narco(i).Qty = ch(cid).narco(i).Qty - Quan
                ch(cid).Coat = ch(cid).Coat - Quan
                If ch(cid).narco(i).Qty = 0 Then
                    ch(cid).narco(i).Price = 0
                End If
            End If
        Next
        Money = Money + (Quan * tmpcost)
        Text2.Text = Text2.Text & "You sold " & Quan & " " & DrugN & " at $" & tmpcost
        Text2.Text = Text2.Text & vbCrLf
        Text2.Text = Text2.Text & "You now have $" & Money
        Text2.Text = Text2.Text & vbCrLf
    ElseIf com = "dc" Then
        cid = Val(InputBox("Character ID:", "Drug Information"))
        Call DrugInfoC(cid)
    ElseIf com = "dt" Then
        Call DrugInfoT
    ElseIf com = "bd" Then
        DrugN = gtemp1
        Quan = gtemp2
        cid = gtemp3
        For i = 1 To 21
            If i = 21 Then
                Y = MsgBox("Drug does not exist", vbInformation, "Error")
                Exit Function
            End If
            If Trim(ch(cid).narco(i).name) = Trim(DrugN) Then Exit For
        Next
        Open "Drugs.txt" For Input As #1
        j = 0
        Do Until EOF(1)
            j = j + 1
            mode = 1
            Line Input #1, X
            If j = 1 Then nt = Val(X)
            For i = 1 To Len(X)
                If mode = 1 And Mid(X, i, 1) <> "," Then
                    tmptext = tmptext & Mid(X, i, 1)
                ElseIf mode = 1 And Mid(X, i, 1) = "," Then
                    tmpdrug(j).name = tmptext
                    tmptext = ""
                    mode = 2
                ElseIf mode = 2 Then
                    tmptext = tmptext & Mid(X, i, 1)
                    tmpdrug(j).Price = Val(tmptext)
                End If
            Next
            tmptext = ""
        Loop
        Close #1
        nt = nt + 1
        For i = 1 To nt
            If DrugN = Trim(tmpdrug(i).name) Then
                tmpcost = tmpdrug(i).Price
                tmptext = Trim(tmpdrug(i).name)
            End If
        Next
        For i = 1 To nt
            If dOffset(i).dName = DrugN Then
                tmpcost = tmpcost + dOffset(i).Price
                If tmpcost * Quan > Money Then
                    Y = MsgBox("You don't have enough money", vbInformation, "Can't Buy")
                    Exit Function
                End If
            End If
        Next
        For i = 1 To nt
            If ch(cid).narco(i).name = DrugN Then
                ch(cid).narco(i).Qty = ch(cid).narco(i).Qty + Quan
                ch(cid).Coat = ch(cid).Coat + Quan
                ch(cid).narco(i).Price = tmpcost
            End If
        Next
        Money = Money - (Quan * tmpcost)
        Text2.Text = Text2.Text & "You bought " & Quan & " " & DrugN & " at $"
        Text2.Text = Text2.Text & tmpcost & "."
        Text2.Text = Text2.Text & vbCrLf
        Text2.Text = Text2.Text & "You now have $" & Money
        Text2.Text = Text2.Text & vbCrLf
    ElseIf com = "et" Then
        Call SetPriceOffset
        Debt = Debt + (Debt * lIntrest)
        Savings = Savings + (Savings * bIntrest)
        DayNum = DayNum + 1
    ElseIf com = "dd" Then
        DrugN = gtemp1
        Quan = gtemp2
        cid = gtemp3
        If Quan = 0 Then
            Y = MsgBox("Impossible to sell nothing", vbInformation, "Error")
            Exit Function
        End If
        Text2.Text = Text2.Text & vbCrLf
        For i = 1 To 21
            If i = 21 Then
                Y = MsgBox("Drug does not exist", vbInformation, "Error")
                Exit Function
            End If
            If Trim(ch(cid).narco(i).name) = Trim(DrugN) Then Exit For
        Next
        For i = 1 To 20
            If ch(cid).narco(i).name = DrugN Then
                If ch(cid).narco(i).Qty < Quan Then
                    Y = MsgBox("You don't have that much " & DrugN & "!", vbInformation, "Error")
                    Exit Function
                End If
            End If
        Next
        Open "Drugs.txt" For Input As #1
        j = 0
        Do Until EOF(1)
            j = j + 1
            mode = 1
            Line Input #1, X
            If j = 1 Then nt = Val(X)
            For i = 1 To Len(X)
                If mode = 1 And Mid(X, i, 1) <> "," Then
                    tmptext = tmptext & Mid(X, i, 1)
                ElseIf mode = 1 And Mid(X, i, 1) = "," Then
                    tmpdrug(j).name = tmptext
                    tmptext = ""
                    mode = 2
                ElseIf mode = 2 Then
                    tmptext = tmptext & Mid(X, i, 1)
                    tmpdrug(j).Price = Val(tmptext)
                End If
            Next
            tmptext = ""
        Loop
        Close #1
        nt = nt + 1
        For i = 1 To nt
            If DrugN = Trim(tmpdrug(i).name) Then
                tmpcost = tmpdrug(i).Price
                tmptext = Trim(tmpdrug(i).name)
            End If
        Next
        For i = 1 To nt
            If dOffset(i).dName = DrugN Then
                tmpcost = tmpcost + dOffset(i).Price
            End If
        Next
        For i = 1 To nt
            If ch(cid).narco(i).name = DrugN Then
                ch(cid).narco(i).Qty = ch(cid).narco(i).Qty - Quan
                ch(cid).Coat = ch(cid).Coat - Quan
                If ch(cid).narco(i).Qty = 0 Then
                    ch(cid).narco(i).Price = 0
                End If
            End If
        Next
    End If
End Function
Sub DrugInfoT()
    On Error Resume Next
    Dim tmpdrug(20) As Drug
    Open "Drugs.txt" For Input As #1
    j = 0
    Do Until EOF(1)
        j = j + 1
        mode = 1
        Line Input #1, X
        If j = 1 Then nt = Val(X)
        For i = 1 To Len(X)
            If mode = 1 And Mid(X, i, 1) <> "," Then
                tmptext = tmptext & Mid(X, i, 1)
            ElseIf mode = 1 And Mid(X, i, 1) = "," Then
                tmpdrug(j).name = tmptext
                tmptext = ""
                mode = 2
            ElseIf mode = 2 Then
                tmptext = tmptext & Mid(X, i, 1)
                tmpdrug(j).Price = Val(tmptext)
                dOffset(j).dName = tmpdrug(j).name
            End If
        Next
        tmptext = ""
    Loop
    Close #1
    For i = 1 To nt + 1
        If tmpdrug(i).name = dOffset(i).dName And tmpdrug(i).name <> "" Then
            Text2.Text = Text2.Text & dOffset(i).dName & "; Price: $"
            Text2.Text = Text2.Text & dOffset(i).Price + tmpdrug(i).Price
            Text2.Text = Text2.Text & vbCrLf
        End If
    Next
End Sub
Sub SetPriceOffset()
    On Error Resume Next
    Dim tmpdrug(20) As Drug, tmprand As Integer, tmpi As Integer
    Randomize Timer
    Open "Drugs.txt" For Input As #1
    j = 0
    Do Until EOF(1)
        j = j + 1
        mode = 1
        Line Input #1, X
        If j = 1 Then nt = Val(X)
        For i = 1 To Len(X)
            If mode = 1 And Mid(X, i, 1) <> "," Then
                tmptext = tmptext & Mid(X, i, 1)
            ElseIf mode = 1 And Mid(X, i, 1) = "," Then
                tmpdrug(j).name = tmptext
                tmptext = ""
                mode = 2
            ElseIf mode = 2 Then
                tmptext = tmptext & Mid(X, i, 1)
                tmpdrug(j).Price = Val(tmptext)
                dOffset(j).dName = tmpdrug(j).name
            End If
        Next
        dOffset(j).dName = tmpdrug(j).name
        dOffset(j).Price = Int((tmpdrug(j).Price - 2 + 1) * Rnd + 1)
        If tmpdrug(j).name = "" Then GoTo afterben
        tmprand = Int((150 - 1 + 1) * Rnd + 1)
        If tmprand = 2 Then
            Y = MsgBox("Pigs just did a big " & tmpdrug(j).name & " bust. Prices are sky-high!", vbInformation, "Drug Bust")
            dOffset(j).dName = tmpdrug(j).name
            dOffset(j).Price = Int(((tmpdrug(j).Price * 9) - 2 + 1) * Rnd + 1)
        ElseIf tmprand = 3 Then
            Y = MsgBox("You found some " & tmpdrug(j).name & " laying in an alley.", vbInformation, "Lucky Find")
            For h = 1 To nt
                If ch(1).narco(h).name = tmpdrug(j).name Then
                    tmpi = Int((100 - 15 + 1) * Rnd + 15)
                    ch(1).narco(h).Qty = ch(1).narco(h).Qty + tmpi
                    ch(1).Coat = ch(1).Coat + tmpi
                End If
            Next
        ElseIf tmprand = 4 Then
            Y = MsgBox("You found some cash on the ground!", vbInformation, "Lucky Find")
            Money = Money + Int((3000 - 100 + 1) * Rnd + 100)
        ElseIf tmprand = 5 Then
            Y = MsgBox("You meet an old friend, he gives you some " & tmpdrug(j).name, vbInformation, "Luck")
            For h = 1 To nt
                If ch(1).narco(h).name = tmpdrug(j).name Then
                    tmpi = Int((100 - 15 + 1) * Rnd + 15)
                    ch(1).narco(h).Qty = ch(1).narco(h).Qty + tmpi
                    ch(1).Coat = ch(1).Coat + tmpi
                End If
            Next
        ElseIf tmprand = 6 Then
            Y = MsgBox("You find some " & tmpdrug(j).name & " on a dead guy in an alley.", vbInformation, "Lucky Find")
            For h = 1 To nt
                If ch(1).narco(h).name = tmpdrug(j).name Then
                    tmpi = Int((100 - 15 + 1) * Rnd + 15)
                    ch(1).narco(h).Qty = ch(1).narco(h).Qty + tmpi
                    ch(1).Coat = ch(1).Coat + tmpi
                End If
            Next
        ElseIf tmprand = 7 Then
            For h = 1 To nt
                If ch(1).narco(h).name = tmpdrug(j).name Then
                    If ch(1).narco(h).Qty = 0 Then
                        Close #1
                        Exit Sub
                    End If
                End If
            Next
            Y = MsgBox("Officer Hemmingstead catches you and confiscates some " & tmpdrug(j).name, vbInformation, "Busted!")
            For h = 1 To nt
                If ch(1).narco(h).name = tmpdrug(j).name Then
                    tmpi = Int((100 - 15 + 1) * Rnd + 15)
                    ch(1).narco(h).Qty = ch(1).narco(h).Qty - tmpi
                    ch(1).Coat = ch(1).Coat - tmpi
                End If
            Next
        End If
afterben:
        tmptext = ""
    Loop
    Close #1
End Sub
Sub DrugInfoC(cid As Integer)
    On Error Resume Next
    For i = 1 To 20
        If ch(cid).narco(i).name <> "" Then
            Text2.Text = Text2.Text & ch(cid).narco(i).name
            Text2.Text = Text2.Text & "; #: "
            Text2.Text = Text2.Text & ch(cid).narco(i).Qty
            Text2.Text = Text2.Text & "; Orig. $: "
            Text2.Text = Text2.Text & "$" & ch(cid).narco(i).Price
            Text2.Text = Text2.Text & vbCrLf
        End If
    Next
End Sub
Sub MoveChar(X As Integer, Y As Integer, cid As Integer)
    On Error Resume Next
    Dim temp1 As Integer, temp2 As Integer
    If Sqr((X - ch(cid).X) ^ 2 + (Y - ch(cid).Y) ^ 2) > ch(cid).Move Then
        Text2.Text = Text2.Text & "Distance too Great"
        Text2.Text = Text2.Text & vbCrLf
    End If
    If map(X, Y).Occupied = True Then
        If map(X, Y).nOcc > 0 Then
            For i = 1 To 100
                If ch(map(X, Y).ocID(i)).Nationality <> ch(cid).Nationality And ch(map(X, Y).ocID(i)).Health <> 0 Then
                    temp1 = map(X, Y).ocID(i)
                    temp2 = ch(cid).ID
                    Call Fight(temp1, temp2)
                End If
            Next
        End If
    Else
        map(ch(cid).X, ch(cid).Y).nOcc = map(ch(cid).X, ch(cid).Y).nOcc - 1
        If map(ch(cid).X, ch(cid).Y).nOcc = 0 Then
            map(ch(cid).X, ch(cid).Y).Occupied = False
        End If
        ch(cid).X = X
        ch(cid).Y = Y
        map(X, Y).Occupied = True
        map(X, Y).nOcc = map(X, Y).nOcc + 1
        map(X, Y).ocID(cid) = 0
        Text2.Text = Text2.Text & Trim(ch(cid).Nationality) & " " & Trim(ch(cid).chName) & _
                     " #" & cid & " moved to " & X & ", " & Y
        Text2.Text = Text2.Text & vbCrLf
        Call SetPriceOffset
    End If
End Sub
Sub Fight(id1 As Integer, id2 As Integer)
    On Error Resume Next
    Randomize Timer
    MsgBox id1
    Do While 1
        If Int((ch(id1).Attack - 0 + 1) * Rnd + 0) > Int((ch(id2).Attack - 0 + 1) * Rnd + 0) Then
            ch(id2).Health = ch(id2).Health - ch(id1).Attack
        Else
            ch(id1).Health = ch(id1).Health - ch(id2).Attack
        End If
        If ch(id1).Health <= 0 Then
            Text2.Text = Text2.Text & Trim(ch(id2).Nationality) & " " & Trim(ch(id2).chName) & _
                         " #" & id2 & " killed " & Trim(ch(id1).Nationality) & _
                         " " & Trim(ch(id1).chName) & " #" & id1
            Text2.Text = Text2.Text & vbCrLf
            map(ch(id1).X, ch(id1).Y).ocID(id1) = 0
            map(ch(id1).X, ch(id1).Y).nOcc = map(ch(id1).X, ch(id1).Y).nOcc - 1
            If map(ch(id1).X, ch(id1).Y).nOcc = 0 Then
                map(ch(id1).X, ch(id1).Y).Occupied = False
            End If
            Exit Do
        ElseIf ch(id2).Health <= 0 Then
            Text2.Text = Text2.Text & Trim(ch(id1).Nationality) & " " & Trim(ch(id1).chName) & _
                         " #" & id1 & " Killed " & Trim(ch(id2).Nationality) & _
                         " " & Trim(ch(id2).chName) & " #" & id2
            Text2.Text = Text2.Text & vbCrLf
            map(ch(id2).X, ch(id2).Y).ocID(id2) = 0
            map(ch(id2).X, ch(id2).Y).nOcc = map(ch(id2).X, ch(id2).Y).nOcc - 1
            If map(ch(id2).X, ch(id2).Y).nOcc = 0 Then
                map(ch(id2).X, ch(id2).Y).Occupied = False
            End If
            Exit Do
        End If
    Loop
End Sub
Function DispChInfo(cid As Integer)
    On Error Resume Next
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "Character Information:"
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "ID: " & cid
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "Type: " & ch(cid).chName
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "Owner: " & ch(cid).Nationality
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "Health: " & ch(cid).Health
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "War: " & ch(cid).Attack
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "Move: " & ch(cid).Move
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "Location: " & ch(cid).X & ", " & ch(cid).Y
    Text2.Text = Text2.Text & vbCrLf
End Function
Sub DispTiInfo(X As Integer, Y As Integer)
    On Error Resume Next
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "Tile Information:"
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "Location: " & X & ", " & Y
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & "Occupied: " & map(X, Y).Occupied
    Text2.Text = Text2.Text & vbCrLf
    If map(X, Y).Occupied = True Then
        Text2.Text = Text2.Text & "Occupied By:"
        Text2.Text = Text2.Text & vbCrLf
        For i = 1 To UBound(map(X, Y).ocID)
            If map(X, Y).ocID(i) <> 0 Then
                Text2.Text = Text2.Text & Trim(ch(i).Nationality) & " " & Trim(ch(i).chName) & ", ID# " & i
                Text2.Text = Text2.Text & vbCrLf
            End If
        Next
    End If
End Sub
Sub SaveGame(BaseName As String)
    On Error Resume Next
    For i = 1 To 5
        ga.Chars(i) = ch(i)
    Next
    For i = 1 To 20
        ga.narco(i) = ch(1).narco(i)
    Next
    ga.Nationality = gNat
    ga.plName = doName
    ga.Money = Money
    ga.nChars = nChars
    ga.Debt = Debt
    ga.Savings = Savings
    ga.gDay = DayNum
    ga.gCity = pCity
    ga.gLoc = CurLoc
    Open BaseName For Random As #1 Len = Len(ga)
    Put #1, 1, ga
    Close #1
End Sub
Sub LoadGame(BaseName As String)
    On Error Resume Next
    Dim nTile As Integer
    Open BaseName For Random As #1 Len = Len(ga)
    Get #1, 1, ga
    Close #1
    For i = 1 To 20
        ch(1).narco(i) = ga.narco(i)
        ch(i) = ga.Chars(i)
    Next
    gNat = ga.Nationality
    doName = ga.plName
    Money = ga.Money
    Debt = ga.Debt
    Savings = ga.Savings
    nChars = ga.nChars
    DayNum = ga.gDay
    pCity = Trim(ga.gCity)
    CurLoc = Trim(ga.gLoc)
End Sub
Function CurToUS(pmoney As Currency) As Currency
    On Error Resume Next
    Dim tmp As String, newm As String
    tmp = Trim(Str(pmoney))
    For i = 1 To Len(tmp)
        If Mid(tmp, i, 1) = "." Then
            newm = newm & "."
            newm = newm & Mid(tmp, i + 1, 1)
            If Mid(tmp, i + 2, 1) = "" Then
                newm = newm + "0"
            Else
                newm = newm & Mid(tmp, i + 2, 1)
            End If
            Exit For
        Else
            newm = newm & Mid(tmp, i, 1)
        End If
    Next
    CurToUS = Val(Trim(newm))
End Function


