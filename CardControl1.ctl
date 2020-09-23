VERSION 5.00
Begin VB.UserControl CardDeck 
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   MaskPicture     =   "CardControl1.ctx":0000
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   523
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      Picture         =   "CardControl1.ctx":04CA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   58
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   57
      Left            =   3720
      Picture         =   "CardControl1.ctx":090C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   57
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   56
      Left            =   3120
      Picture         =   "CardControl1.ctx":170E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   56
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   55
      Left            =   2520
      Picture         =   "CardControl1.ctx":2510
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   55
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   54
      Left            =   1920
      Picture         =   "CardControl1.ctx":3312
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   54
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   53
      Left            =   1320
      Picture         =   "CardControl1.ctx":4114
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   53
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   52
      Left            =   720
      Picture         =   "CardControl1.ctx":4F16
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   52
      Tag             =   "10"
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   51
      Left            =   120
      Picture         =   "CardControl1.ctx":5D18
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   51
      Tag             =   "10"
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   50
      Left            =   5520
      Picture         =   "CardControl1.ctx":6B1A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   50
      Tag             =   "10"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   49
      Left            =   4920
      Picture         =   "CardControl1.ctx":791C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   49
      Tag             =   "10"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   48
      Left            =   4320
      Picture         =   "CardControl1.ctx":7DE6
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   48
      Tag             =   "9"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   47
      Left            =   3720
      Picture         =   "CardControl1.ctx":82B0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   47
      Tag             =   "8"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   46
      Left            =   3120
      Picture         =   "CardControl1.ctx":877A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   46
      Tag             =   "7"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   45
      Left            =   2520
      Picture         =   "CardControl1.ctx":8C44
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   45
      Tag             =   "6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   44
      Left            =   1920
      Picture         =   "CardControl1.ctx":910E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   44
      Tag             =   "5"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   43
      Left            =   1320
      Picture         =   "CardControl1.ctx":95D8
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   43
      Tag             =   "4"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   42
      Left            =   720
      Picture         =   "CardControl1.ctx":9AA2
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   42
      Tag             =   "3"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   41
      Left            =   120
      Picture         =   "CardControl1.ctx":9F6C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   41
      Tag             =   "2"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   40
      Left            =   5520
      Picture         =   "CardControl1.ctx":A436
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   40
      Tag             =   "1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   39
      Left            =   4920
      Picture         =   "CardControl1.ctx":B238
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   39
      Tag             =   "10"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   38
      Left            =   4320
      Picture         =   "CardControl1.ctx":C03A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   38
      Tag             =   "10"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   37
      Left            =   3720
      Picture         =   "CardControl1.ctx":CE3C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   37
      Tag             =   "10"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   36
      Left            =   3120
      Picture         =   "CardControl1.ctx":DC3E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   36
      Tag             =   "10"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   35
      Left            =   2520
      Picture         =   "CardControl1.ctx":E108
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   35
      Tag             =   "9"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   34
      Left            =   1920
      Picture         =   "CardControl1.ctx":E5D2
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   34
      Tag             =   "8"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   33
      Left            =   1320
      Picture         =   "CardControl1.ctx":EA9C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   33
      Tag             =   "7"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   32
      Left            =   720
      Picture         =   "CardControl1.ctx":EF66
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   32
      Tag             =   "6"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   31
      Left            =   120
      Picture         =   "CardControl1.ctx":F430
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   31
      Tag             =   "5"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   30
      Left            =   5520
      Picture         =   "CardControl1.ctx":F8FA
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   30
      Tag             =   "4"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   29
      Left            =   4920
      Picture         =   "CardControl1.ctx":FDC4
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   29
      Tag             =   "3"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   28
      Left            =   4320
      Picture         =   "CardControl1.ctx":1028E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   28
      Tag             =   "2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   27
      Left            =   3720
      Picture         =   "CardControl1.ctx":10758
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   27
      Tag             =   "1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   26
      Left            =   3120
      Picture         =   "CardControl1.ctx":10C22
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   26
      Tag             =   "10"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   25
      Left            =   2520
      Picture         =   "CardControl1.ctx":11A24
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   25
      Tag             =   "10"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   24
      Left            =   1920
      Picture         =   "CardControl1.ctx":12826
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   24
      Tag             =   "10"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   23
      Left            =   1320
      Picture         =   "CardControl1.ctx":13628
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   23
      Tag             =   "10"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   22
      Left            =   720
      Picture         =   "CardControl1.ctx":13AF2
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   22
      Tag             =   "9"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   21
      Left            =   120
      Picture         =   "CardControl1.ctx":13FBC
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   21
      Tag             =   "8"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   20
      Left            =   5520
      Picture         =   "CardControl1.ctx":14486
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   20
      Tag             =   "7"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   19
      Left            =   4920
      Picture         =   "CardControl1.ctx":14950
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   19
      Tag             =   "6"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   18
      Left            =   4320
      Picture         =   "CardControl1.ctx":14E1A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   18
      Tag             =   "5"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   17
      Left            =   3720
      Picture         =   "CardControl1.ctx":152E4
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   17
      Tag             =   "4"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   16
      Left            =   3120
      Picture         =   "CardControl1.ctx":157AE
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   16
      Tag             =   "3"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   15
      Left            =   2520
      Picture         =   "CardControl1.ctx":15C78
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   15
      Tag             =   "2"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   14
      Left            =   1920
      Picture         =   "CardControl1.ctx":16142
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   14
      Tag             =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   13
      Left            =   1320
      Picture         =   "CardControl1.ctx":1660C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   13
      Tag             =   "10"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   12
      Left            =   720
      Picture         =   "CardControl1.ctx":1740E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   12
      Tag             =   "10"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   11
      Left            =   120
      Picture         =   "CardControl1.ctx":18210
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   11
      Tag             =   "10"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   10
      Left            =   5520
      Picture         =   "CardControl1.ctx":19012
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      Tag             =   "10"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   9
      Left            =   4920
      Picture         =   "CardControl1.ctx":194DC
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   9
      Tag             =   "9"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   8
      Left            =   4320
      Picture         =   "CardControl1.ctx":199A6
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      Tag             =   "8"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   7
      Left            =   3720
      Picture         =   "CardControl1.ctx":19E70
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      Tag             =   "7"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   6
      Left            =   3120
      Picture         =   "CardControl1.ctx":1A33A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Tag             =   "6"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   5
      Left            =   2520
      Picture         =   "CardControl1.ctx":1A804
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   5
      Tag             =   "5"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   4
      Left            =   1920
      Picture         =   "CardControl1.ctx":1ACCE
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Tag             =   "4"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   3
      Left            =   1320
      Picture         =   "CardControl1.ctx":1B198
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Tag             =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   2
      Left            =   720
      Picture         =   "CardControl1.ctx":1B662
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Tag             =   "2"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   615
      Index           =   1
      Left            =   120
      Picture         =   "CardControl1.ctx":1BB2C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Tag             =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox card 
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "CardDeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public CardImageDC As Long
Public CardImage As Long
Public CN As Integer
Dim CardID(1 To 52)
Dim CardSUIT(1 To 52)
'Event Declarations:





Private Sub UserControl_Initialize()
CardImage = card(1).Picture
CardImageDC = CreateCompatibleDC(UserControl.hdc)
SelectObject CardImageDC, CardImage

'Assign Cards their ID and Suit
Suit = 1
For i = 1 To 52
  x = x + 1
  CardID(i) = x
  CardSUIT(i) = Suit
  If x = 13 Then
    x = 0
    Suit = Suit + 1
  End If
Next i




End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function DrawCard(hdc As Long, CardNumber As Integer, x As Integer, y As Integer) As Variant
Attribute DrawCard.VB_Description = "Draws a card on the screen."
CardImage = card(CardNumber).Picture
SelectObject CardImageDC, CardImage

BitBlt hdc, x, y, 71, 96, CardImageDC, 0, 0, &HCC0020

End Function




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetValue(CardNumber As Integer, ByRef Value As Variant) As Variant
Value = card(CardNumber).Tag

End Function

'MemberInfo=14
Public Function SetValue(CardNumber As Integer, Value As Integer)
card(i).Tag = Value
End Function

'MemberInfo=14
Public Function GetCardID(CardNumber As Integer, ByRef ID As Variant)
ID = CardID(CardNumber)
End Function

'MemberInfo=14
Public Function GetSuit(CardNumber As Integer, ByRef Suit As Variant)
Suit = CardSUIT(CardNumber)
End Function








