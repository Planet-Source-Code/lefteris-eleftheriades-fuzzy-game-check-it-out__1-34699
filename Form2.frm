VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3075
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Reposission Tiles"
      Height          =   195
      Left            =   30
      TabIndex        =   14
      Top             =   2100
      Value           =   1  'Checked
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1590
      TabIndex        =   12
      Top             =   2370
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   510
      TabIndex        =   11
      Top             =   2370
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Background"
      Height          =   1125
      Left            =   0
      TabIndex        =   6
      Top             =   900
      Width           =   3075
      Begin VB.CommandButton Command3 
         Caption         =   "ÿ ÿ"
         Height          =   285
         Left            =   2670
         TabIndex        =   13
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   870
         TabIndex        =   10
         Top             =   660
         Width           =   1725
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Picture"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   690
         Width           =   825
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Color"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   270
         Value           =   -1  'True
         Width           =   705
      End
      Begin Fuzzy.ColorCombo ColorCombo1 
         Height          =   360
         Left            =   870
         TabIndex        =   7
         Top             =   210
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Difficulty"
      Height          =   885
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Width           =   1515
      Begin MSComDlg.CommonDialog CommDlg 
         Left            =   540
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select Background Image"
         FileName        =   "Image files|*.bmp;*.wmf;*.gif;*.jpg;*.jpeg;*.pcx;*.ico;*.tif"
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Diffucult"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   510
         Width           =   945
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Easy"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solid Blocks"
      Height          =   885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1515
      Begin VB.OptionButton Option1 
         Caption         =   "Iron"
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   2
         Top             =   510
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Brown Brick"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim SB%, DL%, BCP%
   If Option1(0).Value Then
      SB = 0
   Else
      SB = 1
   End If
   If Option2(0).Value Then
      DL = 0
   Else
      DL = 1
   End If
   If Option3(0).Value Then
      BCP = 0
   Else
      BCP = 1
   End If
   SaveSetting App.Title, "BrickType", "Value", SB
   SaveSetting App.Title, "Difficulty", "Value", DL
   SaveSetting App.Title, "Background", "Value", BCP
   SaveSetting App.Title, "Background", "Color", ColorCombo1.Color
   SaveSetting App.Title, "Background", "Picture", Text1.Text
   SaveSetting App.Title, "Reposissioning", "Value", Check1.Value
   SaveSetting App.Title, "Registered", "Value", 1
   DoEvents
   Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
  CommDlg.FileName = ""
  CommDlg.ShowOpen
  Text1.Text = CommDlg.FileName
End Sub

Private Sub Form_Load()
   If GetSetting(App.Title, "BrickType", "Value", 0) = 0 Then
      Option1(0).Value = True
   Else
      Option1(1).Value = True
   End If
   
   If GetSetting(App.Title, "Difficulty", "Value", 0) = 0 Then
      Option2(0).Value = True
   Else
      Option2(1).Value = True
   End If
   
   If GetSetting(App.Title, "Background", "Value", 0) = 0 Then
      Option3(0).Value = True
   Else
      Option3(1).Value = True
   End If
   
   ColorCombo1.Color = GetSetting(App.Title, "Background", "Color", 0)
   Text1.Text = GetSetting(App.Title, "Background", "Picture", "")
   
   Check1.Value = GetSetting(App.Title, "Reposissioning", "Value", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Form1.LoadRegSettings
End Sub
