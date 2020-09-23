VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level Editor"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Random"
      Height          =   315
      Left            =   60
      TabIndex        =   21
      Top             =   2910
      Width           =   1155
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   3
      Left            =   3090
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   20
      Top             =   2220
      Width           =   660
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   2
      Left            =   3090
      Picture         =   "Form1.frx":1302
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   19
      Top             =   1500
      Width           =   660
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   1
      Left            =   3090
      Picture         =   "Form1.frx":2604
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   18
      Top             =   780
      Width           =   660
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   0
      Left            =   3090
      Picture         =   "Form1.frx":3906
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   17
      Top             =   60
      Width           =   660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Level"
      Height          =   345
      Left            =   1680
      TabIndex        =   16
      Top             =   2910
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   15
      Left            =   2190
      Picture         =   "Form1.frx":4C08
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   15
      Top             =   2220
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   14
      Left            =   1470
      Picture         =   "Form1.frx":5F0A
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   14
      Top             =   2220
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   13
      Left            =   750
      Picture         =   "Form1.frx":720C
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   13
      Top             =   2220
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   12
      Left            =   30
      Picture         =   "Form1.frx":850E
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   12
      Top             =   2220
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   11
      Left            =   2190
      Picture         =   "Form1.frx":9810
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   1500
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   10
      Left            =   1470
      Picture         =   "Form1.frx":AB12
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   10
      Top             =   1500
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   9
      Left            =   750
      Picture         =   "Form1.frx":BE14
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   9
      Top             =   1500
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   8
      Left            =   30
      Picture         =   "Form1.frx":D116
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   8
      Top             =   1500
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   7
      Left            =   2190
      Picture         =   "Form1.frx":E418
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   7
      Top             =   780
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   6
      Left            =   1470
      Picture         =   "Form1.frx":F71A
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   6
      Top             =   780
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   5
      Left            =   750
      Picture         =   "Form1.frx":10A1C
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   780
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   4
      Left            =   30
      Picture         =   "Form1.frx":11D1E
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   780
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   3
      Left            =   2190
      Picture         =   "Form1.frx":13020
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   60
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   2
      Left            =   1470
      Picture         =   "Form1.frx":14322
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   60
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   1
      Left            =   750
      Picture         =   "Form1.frx":15624
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   60
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Index           =   0
      Left            =   30
      Picture         =   "Form1.frx":16926
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   60
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Map(1 To 4, 1 To 4) As Integer

Private Sub Command1_Click()
Dim XX&, YY&
Dim NCF As Boolean
Dim LevelData$
Dim OldLevelData$
Dim CurrentLevelData$
 For YY = 1 To 4
   For XX = 1 To 4
       If Map(XX, YY) = -1 Then
         NCF = True
       End If
   Next XX
 Next YY
 
 If NCF Then
    MsgBox "You haven't finished the level", vbInformation Or vbOKOnly
 Else
    LevelData$ = ""
    
    For YY = 1 To 4
      For XX = 1 To 4
          LevelData$ = LevelData$ & Encode(Map(XX, YY))
      Next XX
          If YY <> 4 Then LevelData$ = LevelData$ & "-"
    Next YY
    
    Open "..\Levels.txt" For Input As #1
       Do While Not EOF(1)
          Line Input #1, CurrentLevelData$
          OldLevelData$ = OldLevelData$ & CurrentLevelData$ & vbCrLf
       Loop
    Close #1
    
    Open "..\Levels.txt" For Output As #1
          Print #1, OldLevelData$ & LevelData$
    Close #1
    
 End If
End Sub
Function Encode(DataL%) As String
   Select Case DataL
       Case 1: Encode = "X"
       Case 0: Encode = "O"
       Case 2: Encode = ""
       Case 3: Encode = "Ä"
   End Select
End Function

Private Sub Command2_Click()
Dim I%, R%
  For I = 0 To 15
      For R = 1 To Round(Rnd * 3) + 1
        Picture1_Click I
      Next R
  Next I
End Sub

Private Sub Form_Load()
  'Empty the map

  Map(1, 1) = -1:  Map(2, 1) = -1:  Map(3, 1) = -1:  Map(4, 1) = -1
  Map(1, 2) = -1:  Map(2, 2) = -1:  Map(3, 2) = -1:  Map(4, 2) = -1
  Map(1, 3) = -1:  Map(2, 3) = -1:  Map(3, 3) = -1:  Map(4, 3) = -1
  Map(1, 4) = -1:  Map(2, 4) = -1:  Map(3, 4) = -1:  Map(4, 4) = -1
End Sub

Private Sub Picture1_Click(Index As Integer)
Dim FakeIndex%
Dim XX%, YY%
FakeIndex = Index
Do While FakeIndex > 3
   FakeIndex = FakeIndex - 4
   YY% = YY% + 1
Loop
   XX = FakeIndex + 1
   YY% = YY% + 1
   'MsgBox XX & "  " & YY
   
   Map(XX, YY) = Map(XX, YY) + 1
   'Loop 3 Images
   If Map(XX, YY) = 4 Then Map(XX, YY) = 0
   Picture1(Index).Picture = Picture2(Map(XX, YY)).Picture

End Sub
