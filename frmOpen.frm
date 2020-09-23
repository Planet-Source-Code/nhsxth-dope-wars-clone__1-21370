VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dopewars"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Your Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    doneOpen = 57
    Me.Hide
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    doName = Text1.Text
    pCity = List1.Text
    doneOpen = 57
    SaveGame ("game")
    Me.Hide
    Form1.Show
End Sub
Private Sub Form_Load()
    Open "Locations.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, X
        If Mid(X, 1, 3) = "CT:" Then List1.AddItem Mid(X, 4, Len(X) - 3)
    Loop
    Close #1
End Sub
