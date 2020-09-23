VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   8880
   ControlBox      =   0   'False
   FillStyle       =   4  'Upward Diagonal
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check4 
      Caption         =   "T"
      Height          =   225
      Left            =   1500
      TabIndex        =   14
      Top             =   7050
      Value           =   1  'Checked
      Width           =   465
   End
   Begin VB.CheckBox Check3 
      Caption         =   "S"
      Height          =   225
      Left            =   1020
      TabIndex        =   13
      Top             =   7050
      Value           =   1  'Checked
      Width           =   465
   End
   Begin VB.CheckBox Check2 
      Caption         =   "O"
      Height          =   225
      Left            =   540
      TabIndex        =   12
      Top             =   7050
      Value           =   1  'Checked
      Width           =   465
   End
   Begin VB.CheckBox Check1 
      Caption         =   "X"
      Height          =   225
      Left            =   60
      TabIndex        =   11
      Top             =   7050
      Value           =   1  'Checked
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7005
      Left            =   30
      ScaleHeight     =   463
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   589
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   8895
      Begin VB.Timer Timer5 
         Interval        =   750
         Left            =   4920
         Top             =   5490
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2745
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":0CCA
         ScaleHeight     =   183
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   382
         TabIndex        =   7
         Top             =   1770
         Visible         =   0   'False
         Width           =   5730
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   40
         Left            =   1680
         Top             =   3420
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   30
         Left            =   4050
         Top             =   5700
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   1
         Left            =   6780
         Picture         =   "Form1.frx":341B0
         ScaleHeight     =   42
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   87
         TabIndex        =   16
         Top             =   6300
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   0
         Left            =   6690
         Picture         =   "Form1.frx":36D42
         ScaleHeight     =   42
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   87
         TabIndex        =   15
         Top             =   6210
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   3270
         Top             =   5370
      End
      Begin VB.PictureBox Block 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   5
         Left            =   1530
         Picture         =   "Form1.frx":36F84
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   10
         Top             =   5190
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5640
         Index           =   0
         Left            =   7050
         Picture         =   "Form1.frx":3710E
         ScaleHeight     =   376
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   1
         Top             =   1020
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5640
         Index           =   1
         Left            =   6870
         Picture         =   "Form1.frx":4AE90
         ScaleHeight     =   376
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   9
         Top             =   1170
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2745
         Index           =   1
         Left            =   180
         Picture         =   "Form1.frx":4C07A
         ScaleHeight     =   183
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   382
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   5730
      End
      Begin VB.PictureBox Block 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   4
         Left            =   390
         Picture         =   "Form1.frx":4E314
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   6
         Top             =   5580
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox Block 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   2
         Left            =   30
         Picture         =   "Form1.frx":4F616
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   5
         Top             =   5220
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox Block 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   1
         Left            =   780
         Picture         =   "Form1.frx":50918
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   4
         Top             =   5970
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox Block 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   3
         Left            =   780
         Picture         =   "Form1.frx":51C1A
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   3
         Top             =   5220
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox Block 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   0
         Left            =   30
         Picture         =   "Form1.frx":52F1C
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   2
         Top             =   5970
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   5460
         Top             =   5370
      End
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   7035
      Left            =   30
      ScaleHeight     =   465
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   587
      TabIndex        =   17
      Top             =   30
      Width           =   8865
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "VAG Round"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   6480
         TabIndex        =   18
         Top             =   690
         Width           =   555
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declairs
Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Types
'(None) See type RECT in a module
'Enumerations
'(None)
'Class references
Dim Fuzzy As New CaracterObject
'External References
Dim FSO As FileSystemObject
'Variables
Dim FuzzyY As Single
Dim BoxX As Single
Dim BoxY As Single
Dim JP As Integer
Dim CurrentBrick As Integer
Dim Win As Boolean
Dim Correct As Boolean
Dim Map(1 To 4, 1 To 4) As Integer
Dim BoxScape(1 To 4, 1 To 4) As RECT
Dim Difficult As Boolean
Dim Repossisioning As Boolean
Sub LoadRegSettings()
   'On Error Resume Next
   If GetSetting(App.Title, "BrickType", "Value", 0) = 0 Then
      Block(4).Picture = LoadPicture(App.Path & "\Normal\stone.bmp")
      Picture4(1).Picture = LoadPicture(App.Path & "\ground2.bmp")
   Else
      Block(4).Picture = LoadPicture(App.Path & "\Normal\iron.bmp")
      Picture4(1).Picture = LoadPicture(App.Path & "\ground.bmp")
   End If
   
   Difficult = (GetSetting(App.Title, "Difficulty", "Value", 0) = 1)
    
   If GetSetting(App.Title, "Background", "Value", 0) = 0 Then
      Picture1.Picture = LoadPicture("")
      Picture1.BackColor = GetSetting(App.Title, "Background", "Color", 0)
   Else
      Picture1.Picture = LoadPicture(GetSetting(App.Title, "Background", "Picture", "C:\Windows\Gold Weave.bmp"))
   End If
   Repossisioning = GetSetting(App.Title, "Reposissioning", "Value", 1)
   
   If GetSetting(App.Title, "Registered", "Value", 0) = 0 Then
     SaveSetting App.Title, "BrickType", "Value", 0
     SaveSetting App.Title, "Difficulty", "Value", 0
     SaveSetting App.Title, "Background", "Value", 0
     SaveSetting App.Title, "Background", "Color", 0
     SaveSetting App.Title, "Background", "Picture", "C:\Windows\clouds.bmp"
     SaveSetting App.Title, "Registered", "Value", 1
     SaveSetting App.Title, "Reposissioning", "Value", 1
   End If
End Sub
Private Sub Form_Load()
  Fuzzy.SpriteDataFile = App.Path & "\Fuzzy.spr"
  
  LoadRegSettings
  
  LoadMap 3
  DrawMap
  Fuzzy.Draw 4, 480, 330, Picture1.hDC, Picture3(0).hDC, Picture3(1).hDC
  FuzzyY = 370
  JP = 3
  BoxX = 0
  Randomize
  CurrentBrick = RandomBlock(False)
  PlayLargeSound App.Path & "\Fuzzy2.mid", MIDI_Sequence, "fuzzybck"

  'The next 108 lines set the locations of each
  'block. This is needed later for a validation check
  BoxScape(1, 1).Left = 42
  BoxScape(1, 1).Right = 81
  BoxScape(1, 1).Top = 252
  BoxScape(1, 1).Bottom = 291
  
  BoxScape(1, 2).Left = 42
  BoxScape(1, 2).Right = 81
  BoxScape(1, 2).Top = 294
  BoxScape(1, 2).Bottom = 333
  
  BoxScape(1, 3).Left = 42
  BoxScape(1, 3).Right = 81
  BoxScape(1, 3).Top = 338
  BoxScape(1, 3).Bottom = 377
  
  BoxScape(1, 4).Left = 42
  BoxScape(1, 4).Right = 81
  BoxScape(1, 4).Top = 378
  BoxScape(1, 4).Bottom = 417
  '----------------------
  BoxScape(2, 1).Left = 84
  BoxScape(2, 1).Right = 123
  BoxScape(2, 1).Top = 252
  BoxScape(2, 1).Bottom = 291
  
  BoxScape(2, 2).Left = 84
  BoxScape(2, 2).Right = 123
  BoxScape(2, 2).Top = 294
  BoxScape(2, 2).Bottom = 333
  
  BoxScape(2, 3).Left = 84
  BoxScape(2, 3).Right = 123
  BoxScape(2, 3).Top = 338
  BoxScape(2, 3).Bottom = 377
  
  BoxScape(2, 4).Left = 84
  BoxScape(2, 4).Right = 123
  BoxScape(2, 4).Top = 378
  BoxScape(2, 4).Bottom = 417
  '----------------------
  BoxScape(3, 1).Left = 126
  BoxScape(3, 1).Right = 165
  BoxScape(3, 1).Top = 252
  BoxScape(3, 1).Bottom = 291
  
  BoxScape(3, 2).Left = 126
  BoxScape(3, 2).Right = 165
  BoxScape(3, 2).Top = 294
  BoxScape(3, 2).Bottom = 333
  
  BoxScape(3, 3).Left = 126
  BoxScape(3, 3).Right = 165
  BoxScape(3, 3).Top = 338
  BoxScape(3, 3).Bottom = 377
  
  BoxScape(3, 4).Left = 126
  BoxScape(3, 4).Right = 165
  BoxScape(3, 4).Top = 378
  BoxScape(3, 4).Bottom = 417
  '----------------------
  BoxScape(4, 1).Left = 168
  BoxScape(4, 1).Right = 207
  BoxScape(4, 1).Top = 252
  BoxScape(4, 1).Bottom = 291
  
  BoxScape(4, 2).Left = 168
  BoxScape(4, 2).Right = 207
  BoxScape(4, 2).Top = 294
  BoxScape(4, 2).Bottom = 333
  
  BoxScape(4, 3).Left = 168
  BoxScape(4, 3).Right = 207
  BoxScape(4, 3).Top = 338
  BoxScape(4, 3).Bottom = 377
  
  BoxScape(4, 4).Left = 168
  BoxScape(4, 4).Right = 207
  BoxScape(4, 4).Top = 378
  BoxScape(4, 4).Bottom = 417
  '----------------------
End Sub
Sub LoadMap(LevelNumber&)
Dim Level$
Open App.Path & "\Levels.txt" For Input Access Read As #1
     For I = 0 To LevelNumber& - 1
         Line Input #1, Level$
     Next I
Close #1
Map(1, 1) = BlockSTN(Mid(Level, 1, 1))
Map(2, 1) = BlockSTN(Mid(Level, 2, 1))
Map(3, 1) = BlockSTN(Mid(Level, 3, 1))
Map(4, 1) = BlockSTN(Mid(Level, 4, 1))

Map(1, 2) = BlockSTN(Mid(Level, 6, 1))
Map(2, 2) = BlockSTN(Mid(Level, 7, 1))
Map(3, 2) = BlockSTN(Mid(Level, 8, 1))
Map(4, 2) = BlockSTN(Mid(Level, 9, 1))

Map(1, 3) = BlockSTN(Mid(Level, 11, 1))
Map(2, 3) = BlockSTN(Mid(Level, 12, 1))
Map(3, 3) = BlockSTN(Mid(Level, 13, 1))
Map(4, 3) = BlockSTN(Mid(Level, 14, 1))

Map(1, 4) = BlockSTN(Mid(Level, 16, 1))
Map(2, 4) = BlockSTN(Mid(Level, 17, 1))
Map(3, 4) = BlockSTN(Mid(Level, 18, 1))
Map(4, 4) = BlockSTN(Mid(Level, 19, 1))
End Sub

Function BlockSTN(DataL As String) As Integer
   Select Case UCase(DataL)
       Case "X": BlockSTN = 0
       Case "O": BlockSTN = 1
       Case "": BlockSTN = 2
       Case "Ä": BlockSTN = 3
   End Select
End Function

Sub DrawMap()
Dim XI%, YI%
Picture1.Cls
For YI = 1 To 4
  For XI = 1 To 4
      'Draw's the map
      DrawTile (XI * (Block(1).ScaleWidth + 2)), (YI * (Block(1).ScaleHeight + 2)) + 210, (Map(XI, YI))
  Next XI
Next YI

  For XI = 0 To 13
      'Draws the top Border
      DrawTile (XI * (Block(1).ScaleWidth + 2)), 0, 4
      'Draws the bottom border
      If XI <> 11 And XI <> 12 Then 'Skip these two tiles
         DrawTile (XI * (Block(1).ScaleWidth + 2)), 420, 4
      End If
  Next XI
      
      'Draw the basement door
      BitBlt Picture1.hDC, 461, 419, Picture4(0).ScaleWidth, Picture4(0).ScaleHeight, Picture4(0).hDC, 0, 0, SRCAND
      BitBlt Picture1.hDC, 461, 419, Picture4(1).ScaleWidth, Picture4(1).ScaleHeight, Picture4(1).hDC, 0, 0, SRCINVERT
  
  For YI = 0 To 9
      'Draws the left Border
      DrawTile 0, (YI * (Block(1).ScaleHeight + 2)), 4
      'Draws the right border
      DrawTile 546, (YI * (Block(1).ScaleHeight + 2)), 4
  Next YI
  
  DrawTile 42, 42, 4
  DrawTile 84, 42, 4
  DrawTile 126, 42, 4
  DrawTile 42, 84, 4
  DrawTile 84, 84, 4
  DrawTile 42, 126, 4
  'Draw Ladder
  BitBlt Picture1.hDC, 471, 40, Picture2(1).ScaleWidth, Picture2(1).ScaleHeight, Picture2(1).hDC, 0, 0, SRCAND
  BitBlt Picture1.hDC, 471, 40, Picture2(0).ScaleWidth, Picture2(0).ScaleHeight, Picture2(0).hDC, 0, 0, SRCINVERT
End Sub

Sub DrawTile(X&, Y&, ID%)
   If ID = -1 Then Exit Sub
   BitBlt Picture1.hDC, X, Y, Block(5).ScaleWidth, Block(5).ScaleHeight, Block(5).hDC, 0, 0, SRCAND
   BitBlt Picture1.hDC, X, Y, Block(ID).ScaleWidth, Block(ID).ScaleHeight, Block(ID).hDC, 0, 0, SRCINVERT
End Sub

Private Sub Picture1_DblClick()
   Me.WindowState = vbMinimized
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Timer2.Enabled = False
    StopLargeSound "fuzzybck"
    StopLargeSound "sfx"
    StopLargeSound "sfx2"
    DoEvents
    Unload Me
    DoEvents
    End
  End If
  If KeyCode = vbKeyF2 Then
      Label1.Caption = 100
      LoadMap Round(Rnd * 10.4)
      DrawMap
      Fuzzy.Draw 4, 480, 370, Picture1.hDC, Picture3(0).hDC, Picture3(1).hDC
      FuzzyY = 369
      JP = 3
      BoxX = 0
      Randomize
      Correct = False
      CurrentBrick = RandomBlock(Difficult)
      Timer1.Enabled = True
      Timer4.Enabled = False
      Timer5.Enabled = True
  End If
    If KeyCode = vbKeyF3 Then
       Form2.Show
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTMOVE, 0
End Sub

Function IsKeyDown(ByVal Key&) As Boolean
  Dim KS As Integer
  'Returns if a key is press at the moment
  'Great for reading key combinations
  'E.g. Diagonal Up/Right
  'if GetKeyState(vbkeyUp) and GetKeyState(vbkeyRight) then ...
  
  KS = GetKeyState(Key&)
  If KS < 0 Then
     IsKeyDown = True
  Else
     IsKeyDown = False
  End If
End Function

Private Sub Picture5_KeyUp(KeyCode As Integer, Shift As Integer)
  Picture1_KeyUp KeyCode, 1
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Picture1_MouseDown Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()
  Dim UpDownMov As Boolean
  Static LSD As Boolean
  Static InThrowModeX As Boolean
  Static InThrowModeY As Boolean
  Static STPT As Integer
  Dim XI%, YI%
  Dim RT&
  
  DoEvents
  If IsKeyDown(vbKeyUp) Then
     'Top of stage limit
     If FuzzyY > 42 Then
         FuzzyY = FuzzyY - 21
         STPT = STPT + 1
         UpDownMov = True
     End If
  End If
  
  If IsKeyDown(vbKeyDown) Then
     'Bottom of stage limit
     If FuzzyY <= 349 Then
       FuzzyY = FuzzyY + 21
       STPT = STPT - 1
       UpDownMov = True
     End If
  End If
  
   If IsKeyDown(vbKeySpace) Then
       If STPT Mod 2 = 0 Then
         BoxX = 420
         BoxY = FuzzyY + 10
         InThrowModeX = True
         InThrowModeY = False
         PlayLargeSound App.Path & "\sounds\jump1.wav", WaveFiles, "sfx"
         Timer3.Enabled = True
       Else
         PlayLargeSound App.Path & "\sounds\oops.wav", WaveFiles, "sfx"
         Timer3.Enabled = True
       End If
  End If
 '_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
 'The Blocks on top of the stage make the ball fall vertically
 'Once hit by the block given to fuzzy. This is because the iron
 'block stops the horizontical movement be the throw of the block by
 'fuzzy, yet it can't stay still because of the gravity
 'Becides those are the rules of the game
 'A check if the block has hit any of which should be performed
 '[_][_][_][_][_][_][_][_][_]
 '[_][_][_][X] <--This iron block
 '[_][_][_]
 '[_][_]
 '[_]
 '[_]
 '[_]
 If BoxX + 20 <= 208 And BoxY + 20 > 42 And BoxY + 20 < 84 And InThrowModeX Then
       BoxX = 168
       InThrowModeX = False
       InThrowModeY = True
       PlayLargeSound App.Path & "\sounds\hit.wav", WaveFiles, "sfx"
       Timer3.Enabled = True
 End If
 '[_][_][_][_][_][_][_][_][_]
 '[_][_][_][_]
 '[_][_][X]<--This iron block
 '[_][_]
 '[_]
 '[_]
 '[_]
 
 If BoxX + 20 <= 168 And BoxY + 20 > 84 And BoxY + 20 < 126 And InThrowModeX Then
       BoxX = 126
       InThrowModeX = False
       InThrowModeY = True
       PlayLargeSound App.Path & "\sounds\hit.wav", WaveFiles, "sfx"
       DoEvents
       Timer3.Enabled = True
 End If
 '[_][_][_][_][_][_][_][_][_]
 '[_][_][_][_]
 '[_][_][_]
 '[_][X]<--This iron block
 '[_]
 '[_]
 '[_]
 If BoxX + 20 <= 126 And BoxY + 20 > 126 And BoxY + 20 < 168 And InThrowModeX Then
       BoxX = 84
       InThrowModeX = False
       InThrowModeY = True
       PlayLargeSound App.Path & "\sounds\hit.wav", WaveFiles, "sfx"
       DoEvents
       Timer3.Enabled = True
 End If
 '[_][_][_][_][_][_][_][_][_]
 '[_][_][_][_]
 '[_][_][_]
 '[_][_]
 '[X] <--This iron block and the rest below
 '[X]
 '[X]
 If BoxX + 20 <= 84 And BoxY + 20 > 168 And InThrowModeX Then
       BoxX = 42
       InThrowModeX = False
       InThrowModeY = True
       PlayLargeSound App.Path & "\sounds\hit.wav", WaveFiles, "sfx"
       DoEvents
       Timer3.Enabled = True
 End If
 'End of top block check
 
 'Handle's the kinisis of the block when thrown
 If InThrowModeX Then
     BoxX = BoxX - 20
 End If
  
 'Handle the gravity
 If InThrowModeY Then
     BoxY = BoxY + 20
 End If
  
 'While the user is pressing the Up Or Down key
 'Show the appropriate animation
 If UpDownMov Then
     If JP = 3 Then
        JP = 4
     Else
        JP = 3
     End If
    DoEvents
    'PlayLargeSound App.Path & "\sounds\walk.wav", WaveFiles, "sfx"
    PlaySoundPart App.Path & "\sounds\walk.wav", 1, 150, WaveFiles, "sfx2"
    'DoEvents
    'Timer3.Enabled = True
 Else
     'If the user doen't press anything
     'Show him blinking his eyes
     'every now or then
     'A little trick so that this
     'blink occures each second run of this loop
     LSD = Not LSD
     If LSD Then
       If JP = 5 Then
          JP = 6
       Else
          JP = 5
       End If
     End If
  End If
  
  'Check For Boundries
  'Chechs if the block thrown, has hit any other blocks
  For XI = 4 To 1 Step -1
    For YI = 1 To 4
      'If the block has hit an other
      If BoxY + 20 >= BoxScape(XI, YI).Top And BoxY + 20 <= BoxScape(XI, YI).Bottom _
      And BoxX + 20 >= BoxScape(XI, YI).Left And BoxX + 20 <= BoxScape(XI, YI).Right Then
         'If bricks are of the same type
         If Map(XI, YI) = CurrentBrick Then
            'Erase the brick from the pile
            Map(XI, YI) = -1
            Correct = True
            DoEvents
            PlayLargeSound App.Path & "\sounds\collect.wav", WaveFiles, "sfx"
            'PlaySoundPart App.Path & "\sounds\collect.wav", 1, 100, WaveFiles, "sfx2"
            DoEvents
            Timer3.Enabled = True
         ElseIf Map(XI, YI) <> -1 Then
         'If they do not match and the tile met is visible
            'Stop the trown tile moovement
            InThrowModeY = False
            InThrowModeX = False
            'Reset it's posission
            BoxX = 0
            BoxY = 0
            
            'The "throw" brick changes not as it finds a similar
            'brick, but when it found similar and later hits a different
            'This is done because with one shot two similar bricks
            'to the one trown can be both cleared with one shot
            If Correct Then
               If Repossisioning Then
                  FallingUnpiledTiles
               End If
               CurrentBrick = RandomBlock(Difficult)
            End If
            Correct = False
         End If
      End If
    Next YI
    DoEvents
  Next XI
  'If the brick meets the bottom of the stage
  'Do the same as it would if it had met the wrong brick
  DoEvents
  If BoxY > 390 Then
     InThrowModeY = False
     InThrowModeX = False
     BoxX = 0
     BoxY = 0
     RT& = RandomBlock(Difficult)
     'The RandomBlock() returns -2 only if it can't find any
     'Bricks in the stage. That means the player has won
     If RT& <> -2 Then
        'If the player hasn't won
        If Correct Then
           If Repossisioning Then
              FallingUnpiledTiles
           End If
           CurrentBrick = RandomBlock(Difficult)
        End If
        Correct = False
     Else
        'Notify the player of his victory
         Win = True
         Timer1.Enabled = False
         Timer4.Enabled = True
         Correct = False
         Timer5.Enabled = False
     End If
  End If
  DoEvents
  DrawMap
  
  'Draw the tile thrown
  If BoxX = 0 And BoxY = 0 Then
     DrawTile 450, FuzzyY + 10, CurrentBrick
  Else
     DrawTile (BoxX), (BoxY), CurrentBrick
  End If
  DoEvents
  'Draw fuzzy
  Fuzzy.Draw JP, 480, (FuzzyY), Picture1.hDC, Picture3(0).hDC, Picture3(1).hDC
  'refresh the buffer
  Picture1.Refresh
  DoEvents
  'I use Picture1 as buffer and copy the image to Picture5.
  Picture5.Cls
  BitBlt Picture5.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY
  Picture5.Refresh
  DoEvents
End Sub

Function RandomBlock(AllowCoveredBloock As Boolean) As Integer
Dim HasSquair As Boolean
Dim HasCircle As Boolean
Dim HasTriangle As Boolean
Dim HasX As Boolean
Dim XI%, YI%
Dim TempRand%
Dim TV As Boolean
'Note The check boxes are just for debugging
'and only for viewing the results of this check
 Check1.Value = 0
 Check2.Value = 0
 Check3.Value = 0
 Check4.Value = 0
Select Case AllowCoveredBloock
Case True
'This function may only generate a block with a shape
'Only if it exists in the pile
  For XI = 1 To 4
    For YI = 1 To 4
        'Loop To the whole pile and check
        'if X,Circle,Squair, or triangle exist
        Select Case Map(XI, YI)
             Case 0: HasX = True: Check1.Value = 1
             Case 1: HasCircle = True: Check2.Value = 1
             Case 2: HasSquair = True: Check3.Value = 1
             Case 3: HasTriangle = True: Check4.Value = 1
        End Select
    Next YI
  Next XI
  

  Case False
  'Now  things are harder to calculate
  'Calculate for each row
  For YI = 1 To 4
    For XI = 4 To 1 Step -1
     If Map(XI, YI) <> -1 Then
        Select Case Map(XI, YI)
             Case 0: HasX = True: Check1.Value = 1
             Case 1: HasCircle = True: Check2.Value = 1
             Case 2: HasSquair = True: Check3.Value = 1
             Case 3: HasTriangle = True: Check4.Value = 1
        End Select
        
        Exit For
     End If
    Next XI
  Next YI

  
  For XI = 1 To 4
    For YI = 1 To 4
     If Map(XI, YI) <> -1 Then
        Select Case Map(XI, YI)
             Case 0: HasX = True: Check1.Value = 1
             Case 1: HasCircle = True: Check2.Value = 1
             Case 2: HasSquair = True: Check3.Value = 1
             Case 3: HasTriangle = True: Check4.Value = 1
        End Select
        
        Exit For
     End If
    Next YI
  Next XI
  
End Select
'Now gather all the data and generate the
'number according to the validiation
If HasX Or HasCircle Or HasSquair Or HasTriangle Then
  Do While Not TV
    TempRand% = Round(Rnd * 3)
    Select Case TempRand%
           Case 0: If HasX Then TV = True
           Case 1: If HasCircle Then TV = True
           Case 2: If HasSquair Then TV = True
           Case 3: If HasTriangle Then TV = True
     End Select
  Loop
  RandomBlock = TempRand
Else
  RandomBlock = -2
End If
End Function

Private Sub Timer2_Timer()
  If PersendagePlayedOfAnOpenedSound("fuzzybck") = 100 Then
     StopLargeSound "fuzzybck"
     PlayLargeSound App.Path & "\Fuzzy2.mid", MIDI_Sequence, "fuzzybck"
  End If
End Sub

Private Sub Timer3_Timer()
  If PersendagePlayedOfAnOpenedSound("sfx") = 100 Then
    StopLargeSound "sfx"
    Timer3.Enabled = False
  End If
End Sub

Private Sub Timer4_Timer()
  Static SSS As Integer
  Dim FY As Integer
  Static LastTickCount&
  If Win Then
    If SSS = 1 Then
       SSS = 2
       FY = 320
    Else
       SSS = 1
       FY = 350
    End If
    DrawMap
    Fuzzy.Draw SSS, 480, FY, Picture1.hDC, Picture3(0).hDC, Picture3(1).hDC
    Picture1.Refresh
    
    If GetTickCount - LastTickCount > 510 Then
      LastTickCount = GetTickCount
      PlayLargeSound App.Path & "\sounds\victory.wav", WaveFiles, "sfx"
      Timer3.Enabled = True
    End If
  Else
    FY = 350
    SSS = SSS + 1
    If SSS = 5 Then SSS = 1
    DrawMap
    Fuzzy.Draw SSS + 6, 480, FY, Picture1.hDC, Picture3(0).hDC, Picture3(1).hDC
    Picture1.Refresh
    
    If GetTickCount - LastTickCount > 300 Then
      LastTickCount = GetTickCount
       PlayLargeSound App.Path & "\sounds\nooo.wav", WaveFiles, "sfx"
      Timer3.Enabled = True
    End If
  End If
  DoEvents
  'I use Picture1 as buffer and copy the image to Picture5.
  Picture5.Cls
  BitBlt Picture5.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY
  Picture5.Refresh
  DoEvents
End Sub

Private Sub Timer5_Timer()
    Label1.Caption = Label1.Caption - 1
    If Label1.Caption = 20 Then
       PlayLargeSound App.Path & "\sounds\hurry.wav", WaveFiles, "sfx"
       Timer3.Enabled = True
    End If
    If Label1.Caption = 0 Then
         Win = False
         Timer1.Enabled = False
         Timer4.Enabled = True
         Timer5.Enabled = False
    End If
End Sub
Sub FallingUnpiledTiles()
  Dim XI%, YI%
  For XI = 1 To 4
    For YI = 4 To 2 Step -1
       If Map(XI, YI) = -1 Then
          Map(XI, YI) = Map(XI, YI - 1)
          Map(XI, YI - 1) = -1
       End If
    Next YI
  Next XI
End Sub
