VERSION 5.00
Begin VB.Form frmShark 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Shark"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   Icon            =   "frmShark.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Take Loan"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pay Debt"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmShark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim tmp As Variant
    tmp = Val(InputBox("How much would you like to pay back?", "Pay Debt"))
    If tmp = 0 Then Exit Sub
    If tmp > Money Then
        Y = MsgBox("You don't have that much money!", vbInformation, "Too Much")
        Exit Sub
    End If
    Debt = Debt - tmp
    Money = Money - tmp
    Call RefreshStats
End Sub
Private Sub Command2_Click()
    Dim tmp As Variant
    tmp = Val(InputBox("How much would you like to borrow?", "Borrow Money"))
    If tmp = 0 Then Exit Sub
    Debt = Debt + tmp
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
