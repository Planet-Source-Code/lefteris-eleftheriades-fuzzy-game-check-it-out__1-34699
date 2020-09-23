VERSION 5.00
Begin VB.UserControl ColorCombo 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2115
   ScaleHeight     =   345
   ScaleWidth      =   2115
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1170
      Top             =   -120
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   345
      Left            =   0
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   136
      TabIndex        =   0
      Top             =   0
      Width           =   2100
      Begin VB.CommandButton Command1 
         Caption         =   "‚"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Shape1 
         DrawMode        =   6  'Mask Pen Not
         Height          =   285
         Left            =   0
         Top             =   0
         Width           =   75
      End
   End
End
Attribute VB_Name = "ColorCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Dim XL(26) As Single
Dim ControlID As Long
Dim Colors(26) As Long
'Event Declarations:
Event ValidateColor(Color As Long) 'MappingInfo=Picture1,Picture1,-1,MouseDown
Event ChangeColor(Color As Long)  'MappingInfo=Picture1,Picture1,-1,MouseUp
Event Click() 'MappingInfo=Command1,Command1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

Private Sub Command1_Click()
Dim ColorBoxPossition As RECT
Dim I%
 Call GetWindowRect(UserControl.hwnd, ColorBoxPossition)
 ColorControlfrm.Move ColorBoxPossition.Left * Screen.TwipsPerPixelX, ColorBoxPossition.Bottom * Screen.TwipsPerPixelY
 ColorControlfrm.Visible = True
 ACI = ControlID
 For I% = 0 To 21
  ColorControlfrm.Picture3(I%).BackColor = Colors(I%)
 Next I%
 Timer1.Enabled = True
 RaiseEvent Click
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PreviewsColor As Long
For I% = 0 To 24
 If X > XL(I%) And X < XL(I% + 1) Then
  Shape1.Left = XL(I%)
  PreviewsColor = Command1.BackColor
  Command1.BackColor = Picture1.Point(Shape1.Left + 1, Shape1.Top + 1)
  If PreviewsColor <> Command1.BackColor Then RaiseEvent ValidateColor(Command1.BackColor)
 End If
 Next I%
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then Picture1_MouseDown Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()
Dim OPO As Long
 OPO = GetFocus()
 If OPO <> ColorControlfrm.Picture1.hwnd _
 And OPO <> ColorControlfrm.Picture2.hwnd _
 And OPO <> ColorControlfrm.Picture3(1).hwnd _
 And OPO <> ColorControlfrm.Picture3(2).hwnd _
 And OPO <> ColorControlfrm.Picture3(3).hwnd _
 And OPO <> ColorControlfrm.Picture3(4).hwnd _
 And OPO <> ColorControlfrm.Picture3(5).hwnd _
 And OPO <> ColorControlfrm.Picture3(6).hwnd _
 And OPO <> ColorControlfrm.Picture3(7).hwnd _
 And OPO <> ColorControlfrm.Picture3(8).hwnd _
 And OPO <> ColorControlfrm.Picture3(9).hwnd _
 And OPO <> ColorControlfrm.Picture3(10).hwnd _
 And OPO <> ColorControlfrm.Picture3(11).hwnd _
 And OPO <> ColorControlfrm.Picture3(12).hwnd _
 And OPO <> ColorControlfrm.Picture3(13).hwnd _
 And OPO <> ColorControlfrm.Picture3(14).hwnd _
 And OPO <> ColorControlfrm.Picture3(15).hwnd _
 And OPO <> ColorControlfrm.Picture3(17).hwnd _
 And OPO <> ColorControlfrm.Picture3(18).hwnd _
 And OPO <> ColorControlfrm.Picture3(19).hwnd _
 And OPO <> ColorControlfrm.Picture3(20).hwnd _
 And OPO <> ColorControlfrm.Picture3(21).hwnd _
 And OPO <> ColorControlfrm.Picture5.hwnd _
 Then
  ColorControlfrm.Visible = False
  If ACI = ControlID Then
  For I% = 0 To 21
   Picture1.Line (I% * 5.6 + 1, 0)-(I * 5.6 + 1, 20), ColorControlfrm.Picture3(I%).BackColor
   Colors(I%) = ColorControlfrm.Picture3(I%).BackColor
  Next I%
  End If
  Timer1.Enabled = False
 End If
End Sub

Private Sub UserControl_Initialize()
 Load ColorControlfrm
 Picture1.AutoRedraw = True
 Picture1.DrawWidth = 6
 For I% = 0 To 21
  Picture1.Line (I% * 5.6 + 1, 0)-(I * 5.6 + 1, 20), ColorControlfrm.Picture3(I%).BackColor
  Colors(I%) = ColorControlfrm.Picture3(I%).BackColor
 Next I%
 Picture1.Picture = Picture1.Image
 For I% = 0 To 24
  XL(I%) = Shape1.Width * I%
 Next I%
 Randomize
 ControlID = Int(Rnd * 100)

End Sub

Private Sub UserControl_Resize()
 UserControl.Width = 2110
 UserControl.Height = 355
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Command1,Command1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Command1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Command1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Command1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Command1.BackColor = PropBag.ReadProperty("Color", &H8000000F)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Command1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Color", Command1.BackColor, &H8000000F)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ChangeColor(Command1.BackColor)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Command1,Command1,-1,BackColor
Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    Color = Command1.BackColor
End Property

Public Property Let Color(ByVal New_Color As OLE_COLOR)
    Command1.BackColor() = New_Color
    PropertyChanged "Color"
End Property

