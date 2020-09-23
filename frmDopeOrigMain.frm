VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dopewars"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6825
   Icon            =   "frmDopeOrigMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Drugs"
      Height          =   3615
      Left            =   50
      TabIndex        =   9
      Top             =   3120
      Width           =   6735
      Begin VB.CommandButton Command6 
         Caption         =   "Bank"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Loan Shark"
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Dump"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<--- Sell"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Buy --->"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ListBox List3 
         Height          =   2985
         Left            =   4080
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
      Begin VB.ListBox List2 
         Height          =   2985
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Your Stash:"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Local Market:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Locations"
      Height          =   1455
      Left            =   50
      TabIndex        =   6
      Top             =   1570
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
         Height          =   1035
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Label lblCoat 
      BackStyle       =   0  'Transparent
      Caption         =   "0/500"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Pockets:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Samson"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDay 
      BackStyle       =   0  'Transparent
      Caption         =   "0/31"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Day:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblSavings 
      BackStyle       =   0  'Transparent
      Caption         =   "$0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Savings:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblDebt 
      BackStyle       =   0  'Transparent
      Caption         =   "$0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Debt:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblMoney 
      BackStyle       =   0  'Transparent
      Caption         =   "$0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Money:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1470
      Left            =   45
      Top             =   15
      Width           =   6660
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewItem 
         Caption         =   "New &Game"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim hlist As HighScores
    If Command1.Caption = "Go" Then
        If DayNum >= 31 Then Exit Sub
        If List1.Text = "" Then
            Y = MsgBox("Must choose destination!", vbInformation, "No Destination")
            Exit Sub
        End If
        If List1.Text = CurLoc Then
            Y = MsgBox("Already at that location", vbInformation, "Already There")
            Exit Sub
        End If
        CurLoc = List1.Text
        Call Process("et")
        Call RefreshStats
        Call RefreshLoc
        Call CharDrugInfo
        If DayNum = TotalDays Then
            Y = MsgBox("Last Day!", vbInformation, "Last Day")
            Command1.Caption = "New Game"
        End If
        SaveGame ("game")
    Else
        Call NewGame
    End If
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    Dim mode As Integer, tmptext As String
    Dim name As String, cost As String
    mode = 0
    If List2.Text = "" Then
        Y = MsgBox("Select item to buy.", vbInformation, "Buy Error")
        Exit Sub
    End If
    For i = 1 To Len(List2.Text)
        If mode = 0 Then
            name = name & Mid(List2.Text, i, 1)
        ElseIf mode = 1 Then
            tmptext = tmptext & Mid(List2.Text, i, 1)
        End If
        If Mid(List2.Text, i + 2, 1) = "-" Then
            mode = 2
        ElseIf Mid(List2.Text, i, 1) = "$" Then
            mode = 1
        ElseIf i = Len(List2.Text) Then
            cost = Val(tmptext)
        End If
    Next
    gtemp1 = name
    gtemp2 = Val(InputBox("How many would you like to purchase?", "Buy Drugs"))
    gtemp3 = 1
    Call Process("bd")
    Call CharDrugInfo
    Call RefreshStats
End Sub
Private Sub Command3_Click()
    On Error Resume Next
    Dim mode As Integer, tmptext As String
    Dim name As String, cost As String, nname As String
    mode = 0
    If List3.Text = "" Then
        Y = MsgBox("Select item to sell.", vbInformation, "Sell Error")
        Exit Sub
    End If
    For i = 1 To Len(List3.Text)
        name = name & Mid(List3.Text, i, 1)
        If Mid(List3.Text, i + 1, 1) = ";" Then
            Exit For
        End If
    Next
    gtemp1 = name
    gtemp2 = Val(InputBox("How many would you like to sell?", "Sell Drugs"))
    gtemp3 = 1
    Call Process("sd")
    Call CharDrugInfo
    Call RefreshStats
End Sub
Private Sub Command4_Click()
    Dim mode As Integer, tmptext As String
    Dim name As String, cost As String
    mode = 0
     If List2.Text = "" Then
        Y = MsgBox("Select item to dump.", vbInformation, "Dump Error")
        Exit Sub
    End If
    For i = 1 To Len(List2.Text)
        If mode = 0 Then
            name = name & Mid(List2.Text, i, 1)
        ElseIf mode = 1 Then
            tmptext = tmptext & Mid(List2.Text, i, 1)
        End If
        If Mid(List2.Text, i + 2, 1) = "-" Then
            mode = 2
        ElseIf Mid(List2.Text, i, 1) = "$" Then
            mode = 1
        ElseIf i = Len(List2.Text) Then
            cost = Val(tmptext)
        End If
    Next
    gtemp1 = name
    gtemp2 = Val(InputBox("How many would you like to dump?", "Dump Drugs "))
    gtemp3 = 1
    Call Process("dd")
    Call CharDrugInfo
    Call RefreshStats
End Sub
Private Sub Command5_Click()
    frmShark.Show
End Sub
Private Sub Command6_Click()
    frmBank.Show
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Dim shrinkwrap As String
    TotalDays = 31
    shrinkwrap = GetSetting(App.Title, "General", "Wrap")
    If shrinkwrap = "Torn" Then
        Call LoadGame("game")
        Call SetPriceOffset
        Call RefreshLoc
        Call LoadLocations(pCity)
        Call CharDrugInfo
        Call RefreshStats
    Else
        SaveSetting App.Title, "General", "Wrap", "Torn"
        Call NewGame
    End If
End Sub
Private Sub NewGame()
    On Error Resume Next
    Money = 5000
    Debt = 6000
    Savings = 0
    DayNum = 1
    TotalDays = 31
    doneOpen = 3
    frmOpen.Show
    Do Until doneOpen = 57
        DoEvents
    Loop
    Call Process("cc")
    Call SetPriceOffset
    Call RefreshLoc
    Call LoadLocations(pCity)
    Call CharDrugInfo
    Call RefreshStats
    lIntrest = 0.09
    bIntrest = 0.1
End Sub
Private Sub RefreshLoc()
    On Error Resume Next
    Dim mode As Integer
    Dim tmptext As String, tmplist As String
    Text2.Text = ""
    List2.Clear
    Call DrugInfoT
    Call SaveText2
    mode = 0
    Open "tempd" For Input As #1
    Do Until EOF(1)
        Line Input #1, X
        For i = 1 To Len(X)
            If mode = 0 And Mid(X, i, 1) <> "$" Then
                tmptext = tmptext + Mid(X, i, 1)
            ElseIf mode = 2 Then
                tmptext = tmptext + Mid(X, i, 1)
            End If
            If Mid(X, i + 1, 1) = ";" And mode = 0 Then
                tmplist = tmptext
                tmptext = ""
                mode = 1
            ElseIf Mid(X, i, 1) = "$" Then
                mode = 2
            ElseIf i = Len(X) And mode = 2 Then
                mode = 0
                tmplist = tmplist & " - $" & tmptext
                List2.AddItem tmplist
                tmplist = ""
                tmptext = ""
            End If
        Next
    Loop
    Close #1
End Sub
Private Sub SaveText2()
    On Error Resume Next
    Dim linet As String
    Open "tempd" For Output As #1
    For i = 1 To Len(Text2.Text) - 1
        If Asc(Mid(Text2.Text, i + 1, 1)) = 13 Then
            linet = linet + Mid(Text2.Text, i, 1)
            Print #1, linet
            linet = ""
        ElseIf Asc(Mid(Text2.Text, i + 1, 1)) <> 13 And Asc(Mid(Text2.Text, i, 1)) <> 13 And Asc(Mid(Text2.Text, i, 1)) <> 10 Then
            linet = linet + Mid(Text2.Text, i, 1)
        End If
    Next
    Close #1
End Sub
Private Sub LoadLocations(City As String)
    On Error Resume Next
    Dim addb As Boolean, numl As Integer
    addb = False
    numl = 0
    List1.Clear
    Open "Locations.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, X
        If Mid(X, 1, 3) = "CT:" Then
            If addb = True Then
                Close #1
                Exit Sub
            End If
            If Mid(X, 4, Len(X) - 3) = City Then addb = True
        Else
            If addb = True And X <> "" Then
                List1.AddItem X
                If numl = 0 Then CurLoc = X
                numl = 1
            End If
        End If
    Loop
    Close #1
    Call RefreshStats
End Sub
Private Sub CharDrugInfo()
    On Error Resume Next
    Text2.Text = ""
    List3.Clear
    Call DrugInfoC(1)
    Call SaveText2
    Open "tempd" For Input As #1
    Do Until EOF(1)
        Line Input #1, X
        List3.AddItem X
    Loop
    Close #1
End Sub
Private Sub RefreshStats()
    On Error Resume Next
    lblMoney.Caption = "$" & CurToUS(Money)
    lblDebt.Caption = "$" & CurToUS(Debt)
    lblSavings.Caption = "$" & CurToUS(Savings)
    lblDay.Caption = DayNum & "/" & TotalDays
    lblName.Caption = doName
    lblCoat.Caption = ch(1).Coat & " / 500"
    Me.Caption = "Dopewars - " & CurLoc & " @ " & pCity
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call SaveGame("game")
    End
End Sub
Private Sub mnuBankItem_Click()
    On Error Resume Next
    frmBank.Show
End Sub
Private Sub mnuExitItem_Click()
    On Error Resume Next
    SaveGame ("game")
    End
End Sub
Private Sub mnuGameEndItem_Click()
    Dim hlist As HighScores, i As Integer, j As Integer, nlist As HighScores
    Open "highscores" For Random As #1 Len = Len(hlist)
    Get #1, 1, hlist
    Close #1
    For i = 1 To 10
        If (Money + Savings) - Debt > hlist.score(i) Then
            nlist.score(i) = (Money + Savings) - Debt
            nlist.pname(i) = doName
            For j = i + 1 To 9
                nlist.pname(j) = hlist.pname(j - 1)
                nlist.score(j) = hlist.score(j - 1)
            Next
            Exit For
        Else
            nlist.score(i) = hlist.score(i)
            nlist.pname(i) = hlist.pname(i)
        End If
    Next
    Open "highscores" For Random As #1 Len = Len(nlist)
    Put #1, 1, nlist
    Close #1
    frmFinish.Show
End Sub
Private Sub mnuLoadIttem_Click()
    On Error Resume Next
    Call LoadGame("game")
    Call SetPriceOffset
    Call RefreshLoc
    Call LoadLocations(pCity)
    Call CharDrugInfo
    Call RefreshStats
End Sub
Private Sub mnuLoanSharkItem_Click()
    On Error Resume Next
    frmShark.Show
End Sub
Private Sub mnuHighScoresItem_Click()
    frmHighScores.Show
End Sub
Private Sub mnuNewItem_Click()
    On Error Resume Next
    Me.Hide
    frmOpen.Show
    Call NewGame
    Call SaveGame("game")
End Sub
Private Sub mnuSaveItem_Click()
    On Error Resume Next
    Call SaveGame("game")
    Y = MsgBox("Game Saved", vbInformation, "Game Saved")
End Sub
