VERSION 5.00
Begin VB.UserControl BlackBeauty 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1620
   ScaleHeight     =   480
   ScaleWidth      =   1620
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   675
      Top             =   180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   315
      TabIndex        =   0
      Top             =   225
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   0
      Picture         =   "UserControl1.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2460
   End
End
Attribute VB_Name = "BlackBeauty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As Pointapi) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, ByRef lpPoint As Pointapi) As Integer
 
Private Type Pointapi
   X As Long
   Y As Long
End Type


Dim def_forecolor       As Long
Dim timer_counter       As Long

Event Click()
Event MouseEnter()
Event MouseExit()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Default Property Values:
Const m_def_MouseOverCaptionColor = &HFFC0C0

'Property Variables:
Dim m_MouseOverCaptionColor As OLE_COLOR





Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Timer1 = False Then
     TurnTimerOn
     RaiseEvent MouseEnter
  End If
  
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Timer1 = False Then
     TurnTimerOn
     RaiseEvent MouseEnter
  End If

End Sub

Private Sub TurnTimerOn()
   
   timer_counter = 0
   Timer1.Interval = 70
   Timer1 = True
   '  the turning on of this timer indicates the mouse
   '  has just moved over this control so change the
   '  captions forecolor to the value of MouseOverCaptionColor
   Label1.Forecolor = m_MouseOverCaptionColor
   
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Image1.BorderStyle = 1
  RaiseEvent MouseDown(Button, Shift, X, Y)
  
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Image1.BorderStyle = 1
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub
 

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Image1.BorderStyle = 0
  RaiseEvent MouseUp(Button, Shift, X, Y)
  RaiseEvent Click
  
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Image1.BorderStyle = 0
  RaiseEvent MouseUp(Button, Shift, X, Y)
  RaiseEvent Click
  
End Sub


Private Sub Label1_Change()
  '
  'this event will be raised when the caption
  'or text of the label changes so this is a
  'good place to readjust the labels positioning
  '
  Call centerLabel
  
End Sub

Sub centerLabel()
 '
 'this sub centers the label vertically and
 'horiz over the control
 '
 Dim halfOfUCx  As Long, halfOfLblx As Long
 Dim halfOfUCy  As Long, halfOfLbly As Long
 
 halfOfUCx = (Width * 0.5): halfOfLblx = (Label1.Width * 0.5)
 halfOfUCy = (Height * 0.5): halfOfLbly = (Label1.Height * 0.5)
  
 Label1.Move (halfOfUCx - halfOfLblx), (halfOfUCy - halfOfLbly)
 
End Sub
 
 

Private Sub Timer1_Timer()
  '
  'the first function of this timer is to create a nice
  'flicker effect signaling the mouseenter and give us
  'a little bit of a flash button effect
  If timer_counter < 9 Then
     If timer_counter Mod 2 = 0 Then
        Label1.Forecolor = m_MouseOverCaptionColor
     Else
        Label1.Forecolor = def_forecolor
     End If
  End If
  '
  'the purpose of this timer is to track to see when the
  'mouse leaves the boundries of this control and we can
  'raise the mouseexit event
  '
  Dim pt As Pointapi
  GetCursorPos pt
  '  make the value of pt.x and pt.y relative to this
  '  control and not the screen so the position of 0,0
  '  reflects the upper left corner of this control
  ScreenToClient hwnd, pt
  '  now see if the cursor is within the bounds of this
  Dim pixWid As Long, pixHeight  As Long
  pixWid = (Width / Screen.TwipsPerPixelX)
  pixHeight = (Height / Screen.TwipsPerPixelY)
  '  any one of the four following conditions indicate
  '  the mouse is now outside this controls boundries
  If pt.X < 0 Or _
     pt.X > pixWid Or _
     pt.Y < 0 Or _
     pt.Y > pixHeight Then
     
     Timer1 = False
     Label1.Forecolor = def_forecolor
     RaiseEvent MouseExit
   End If
   
   timer_counter = (timer_counter + 1)
   
End Sub

Private Sub UserControl_Resize()
   
   Image1.Move 0, 0, Width, Height
   Call centerLabel
   '
   'here we set min and max size allowed for the
   'control because if its resized to large or small
   'then the image distorts and looks bad
   '
   If Width < 500 Then Width = 500
   If Height < 300 Then Height = 300
   If Width > 3000 Then Width = 3000
   If Height > 1500 Then Height = 1500

   
End Sub

Private Sub UserControl_Show()
 
  'set certain props once
  Label1.AutoSize = True
  Call UserControl_Resize

End Sub
'caption
Public Property Get Caption() As String
    Caption = Label1.Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property
'Forecolor
Public Property Get Forecolor() As OLE_COLOR
Attribute Forecolor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    Forecolor = Label1.Forecolor
End Property
Public Property Let Forecolor(ByVal New_Forecolor As OLE_COLOR)
    Label1.Forecolor() = New_Forecolor
    PropertyChanged "Forecolor"
End Property
'Font
Public Property Get Font() As Font
    Set Font = Label1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
    '  since the font is changing we need to readjust
    '  the positioning of the label so its centered
    Call centerLabel
End Property
'MouseOverCaptionColor
Public Property Get MouseOverCaptionColor() As OLE_COLOR
    MouseOverCaptionColor = m_MouseOverCaptionColor
End Property
Public Property Let MouseOverCaptionColor(ByVal New_MouseOverCaptionColor As OLE_COLOR)
    m_MouseOverCaptionColor = New_MouseOverCaptionColor
    PropertyChanged "MouseOverCaptionColor"
End Property

 
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MouseOverCaptionColor = m_def_MouseOverCaptionColor
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Label1.Caption = PropBag.ReadProperty("Caption", "Caption")
    Label1.Forecolor = PropBag.ReadProperty("Forecolor", &HFFFFFF)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_MouseOverCaptionColor = PropBag.ReadProperty("MouseOverCaptionColor", m_def_MouseOverCaptionColor)
    
    '  save this value for when mouse exits and
    '  we want forecolor to revert to default
    def_forecolor = Label1.Forecolor
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Caption")
    Call PropBag.WriteProperty("Forecolor", Label1.Forecolor, &HFFFFFF)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseOverCaptionColor", m_MouseOverCaptionColor, m_def_MouseOverCaptionColor)
End Sub


