VERSION 5.00
Begin VB.UserControl ValueBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Shape shpBar 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00FFFFFF&
      Height          =   2730
      Left            =   405
      Top             =   30
      Width           =   180
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   2790
      Left            =   375
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "ValueBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ValueBarOrientation
  vboHorizontal
  vboVertical
End Enum


Public Event Change()


Private piValue As Integer
Private piMax As Integer
Private peOrientation As ValueBarOrientation
Private pbDrag As Boolean

Private Sub UserControl_Initialize()
  peOrientation = vboVertical
  Max = 255
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pbDrag = True
  UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If pbDrag Then
    If peOrientation = vboVertical Then
      piValue = piMax * (UserControl.ScaleHeight - Y + 4 * Screen.TwipsPerPixelY) / UserControl.ScaleHeight
    Else
      piValue = piMax * ((X - 4 * Screen.TwipsPerPixelX) / UserControl.ScaleWidth)
    End If
    If piValue < 0 Then piValue = 0
    If piValue > piMax Then piValue = piMax
    update
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pbDrag = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error Resume Next
  With PropBag
    Me.Orientation = .ReadProperty("Orientation", vboVertical)
    Me.Max = .ReadProperty("Max", 255)
    Me.Value = .ReadProperty("Value", 0)
  End With
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "Orientation", Me.Orientation
    .WriteProperty "Max", Me.Max
    .WriteProperty "Value", Me.Value
  End With
End Sub


Private Sub UserControl_Resize()
  On Error Resume Next
  
  With shpBack
    .Left = 0
    .Top = 0
    .Width = UserControl.ScaleWidth
    .Height = UserControl.ScaleHeight
  End With
  With shpBar
    .Left = 2 * Screen.TwipsPerPixelX
    .Top = 2 * Screen.TwipsPerPixelY
    .Width = UserControl.ScaleWidth - 4 * Screen.TwipsPerPixelX
    .Height = UserControl.ScaleHeight - 4 * Screen.TwipsPerPixelY
  End With
End Sub


Public Property Let Value(ByVal iValue As Integer)
  If iValue > piMax Then iValue = piMax
  If iValue < 0 Then iValue = 0
  piValue = iValue
  update
End Property

Public Property Get Value() As Integer
  Value = piValue
End Property


Public Property Let Max(ByVal iMax As Integer)
  If iMax < piValue Then piValue = iMax
  If iMax <= 0 Then iMax = 255
  piMax = iMax
  Value = piValue  ' redraw
End Property

Public Property Get Max() As Integer
  Max = piMax
End Property


Public Property Let Orientation(ByVal eOrientation As ValueBarOrientation)
  If eOrientation = vboHorizontal Or eOrientation = vboVertical Then peOrientation = eOrientation
  update
End Property

Public Property Get Orientation() As ValueBarOrientation
  Orientation = peOrientation
End Property



Private Sub update()
  If peOrientation = vboHorizontal Then
    With shpBar
      .Width = Int(CSng(piValue / piMax) * (UserControl.ScaleWidth - 4 * Screen.TwipsPerPixelX))
    End With
  Else
    With shpBar
      '.Top = CInt(((piMax - piValue) / piMax) * (UserControl.ScaleHeight - 2 * Screen.TwipsPerPixelY)) + Screen.TwipsPerPixelX
      '.Height = UserControl.ScaleHeight - Screen.TwipsPerPixelY - .Top
      .Height = Int(CSng(piValue / piMax) * (UserControl.ScaleHeight - 4 * Screen.TwipsPerPixelY))
      .Top = UserControl.ScaleHeight - .Height - 2 * Screen.TwipsPerPixelY
      
    End With
  End If
  RaiseEvent Change
End Sub
