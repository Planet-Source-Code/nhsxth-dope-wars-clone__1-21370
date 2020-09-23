VERSION 5.00
Begin VB.Form frmBank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   Icon            =   "frmBank.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Withdraw from Savings"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deposit to Savings"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim tmp As Currency
    tmp = Val(InputBox("Deposit how much", "Savings Deposit"))
    If tmp = 0 Then Exit Sub
    If tmp > Money Then
        Y = MsgBox("You don't have that much money!", vbInformation, "Too Much")
        Exit Sub
    End If
    Savings = Savings + tmp
    Money = Money - tmp
    Call RefreshStats
End Sub
Private Sub Command2_Click()
    Dim tmp As Currency
    tmp = Val(InputBox("Withdraw how much", "Savings Withdrawl"))
    If tmp = 0 Then Exit Sub
    If tmp > Savings Then
        Y = MsgBox("You don't have that much savings!", vbInformation, "Too Much")
        Exit Sub
    End If
    Savings = Savings - tmp
    Money = Money + tmp
    Call RefreshStats
End Sub
Private Sub Command3_Click()
    Me.Hide
End Sub
Private Sub RefreshStats()
    Form1.lblMoney.Caption = "$" & CurToUS(Money)
    Form1.lblDebt.Caption = "$" & CurToUS(Debt)
    Form1.lblSavings.Caption = "$" & CurToUS(Savings)
    Form1.lblDay.Caption = DayNum & "/" & TotalDays
End Sub
