VERSION 5.00
Begin VB.Form ColorControlfrm 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   LinkTopic       =   "Form4"
   ScaleHeight     =   2460
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture6 
      DragMode        =   1  'Automatic
      Height          =   165
      Left            =   1800
      ScaleHeight     =   105
      ScaleWidth      =   165
      TabIndex        =   26
      Top             =   480
      Width           =   225
   End
   Begin VB.PictureBox Picture4 
      DragMode        =   1  'Automatic
      Height          =   165
      Left            =   60
      ScaleHeight     =   105
      ScaleWidth      =   165
      TabIndex        =   18
      Top             =   480
      Width           =   225
   End
   Begin VB.PictureBox Picture5 
      Height          =   165
      Left            =   330
      ScaleHeight     =   105
      ScaleWidth      =   1365
      TabIndex        =   25
      Top             =   480
      Width           =   1425
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000018&
      Height          =   195
      Index           =   21
      Left            =   1860
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   24
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000001&
      Height          =   195
      Index           =   20
      Left            =   1680
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   23
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00404040&
      Height          =   195
      Index           =   19
      Left            =   1500
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   22
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00808080&
      Height          =   195
      Index           =   18
      Left            =   1320
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   21
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00A0A0A0&
      Height          =   195
      Index           =   17
      Left            =   1140
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   20
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      Height          =   195
      Index           =   16
      Left            =   960
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   19
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   15
      Left            =   780
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   17
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   600
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   16
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00AA00FF&
      Height          =   195
      Index           =   13
      Left            =   420
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   15
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF00FF&
      Height          =   195
      Index           =   12
      Left            =   240
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   14
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF00AA&
      Height          =   195
      Index           =   11
      Left            =   60
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   13
      Top             =   270
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF0000&
      Height          =   195
      Index           =   10
      Left            =   1860
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   12
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFAA00&
      Height          =   195
      Index           =   9
      Left            =   1680
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   11
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFF00&
      Height          =   195
      Index           =   8
      Left            =   1500
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   10
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00AAFF00&
      Height          =   195
      Index           =   7
      Left            =   1320
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   9
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FF00&
      Height          =   195
      Index           =   6
      Left            =   1140
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   8
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FFBB&
      Height          =   195
      Index           =   5
      Left            =   960
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   7
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   780
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   6
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000BBFF&
      Height          =   195
      Index           =   3
      Left            =   600
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   5
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000080FF&
      Height          =   195
      Index           =   2
      Left            =   420
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000060FF&
      Height          =   195
      Index           =   1
      Left            =   240
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   3
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   60
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   60
      Width           =   165
   End
   Begin VB.PictureBox Picture2 
      Height          =   1755
      Left            =   60
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   1
      Top             =   660
      Width           =   1755
      Begin VB.Image Image1 
         Height          =   1710
         Left            =   0
         Picture         =   "color.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1755
      Left            =   1830
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      Top             =   660
      Width           =   195
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   -30
      X2              =   2040
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   -30
      X2              =   2100
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   0
      X2              =   2100
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   2
      X1              =   0
      X2              =   2070
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   -15
      X2              =   0
      Y1              =   30
      Y2              =   2460
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      Index           =   0
      X1              =   2055
      X2              =   2055
      Y1              =   60
      Y2              =   2610
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   2040
      X2              =   2040
      Y1              =   60
      Y2              =   2670
   End
End
Attribute VB_Name = "ColorControlfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Sub AnalyzeColor(Color&, Target As Object, Optional Frequency As Single = 2.75, Optional DraWWidth_ As Long = 3, Optional SM As Long = 56)
Dim I%, r&, G&, B&
   Target.DrawWidth = DraWWidth_
   GetPixel Color&, r&, G&, B& 'Analize Color to RGB
   If (r < 0) Or (G < 0) Or (B < 0) Then r& = r& + 1: G& = G& + 1: B& = B& + 1
   For I% = 1 To 20
    Target.Line (-10, SM - I% * Frequency)-(30, SM - I% * Frequency), _
    RGB(r& + ((255 - r&) / 20 * I%), G& + ((255 - G&) / 20 * I%), B& + ((255 - B&) / 20 * I%))
   Next I%
   For I% = 0 To 20
    Target.Line (-10, I% * Frequency + SM)-(30, I% * Frequency + SM), _
    RGB(r& - (r& / 20 * I%), G& - (G& / 20 * I%), B& - (B& / 20 * I%))
   Next I%
End Sub

Private Sub Form_Activate()
 Putfocus Picture1.hwnd
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Col As Long
 If Button = 1 Then
  If (X > Picture2.Width - 50) Or (X < 0) Then Exit Sub
  If (Y > Picture2.Height - 50) Or (Y < 0) Then Exit Sub
  Col = Picture2.Point(X, Y)
  If Col = -1 Then Col = 0
  AnalyzeColor Col, Picture1
  Picture4.BackColor = Col
  Picture5.BackColor = Col
  Picture6.BackColor = Col
 End If
End Sub

Function GetPixel(ByVal Colour&, ByRef red&, ByRef green&, ByRef blue&)
 blue& = Int(Colour& / 65536) ' function to get the blue
 green& = Int((Colour& - (65536 * blue&)) / 256) ' function to get the green
 red& = Colour& - (blue& * 65536) - (green& * 256) ' function to get the red
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Picture1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Col As Long
 If Button = 1 Then
  If (X > Picture1.Width - 50) Or (X < 0) Then Exit Sub
  If (Y > Picture1.Height - 50) Or (Y < 0) Then Exit Sub
  Col = Picture1.Point(X, Y)
  If Col = -1 Then Col = 0
  Picture4.BackColor = Col
  Picture5.BackColor = Col
  Picture6.BackColor = Col
 End If
End Sub

Private Sub Picture3_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
 If UCase(Source.Name) = "PICTURE4" Or UCase(Source.Name) = "PICTURE6" Then
  Picture3(Index).BackColor = Picture4.BackColor
 End If
End Sub
