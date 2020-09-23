VERSION 5.00
Object = "{FE59C584-0383-11D6-9D0B-BC0905ED0330}#26.0#0"; "CardDeck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Some Cards"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin CardDeckv1.CardDeck CardDeck1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMain.frx":0000
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Dont forget to set its visibility property to 'false'"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer       'MUST DECLARE CARD NUMBER AS INTEGER!!
Dim x As Integer       'Well, must declare all variables..
For i = 1 To 52
  CardDeck1.DrawCard Me.hDC, i, x, 100
  x = x + 20
Next i

End Sub
