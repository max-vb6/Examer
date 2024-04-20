VERSION 5.00
Begin VB.UserControl ucBtn 
   BackColor       =   &H00FA8C5A&
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   855
   ScaleWidth      =   1695
   Tag             =   "0"
   ToolboxBitmap   =   "ucBtn.ctx":0000
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F5F5F5&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1530
   End
   Begin VB.Shape shpBrd 
      BorderColor     =   &H00D27800&
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "ucBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim IsDft As Boolean
Event Click()

Private Sub lblCap_Click()
    RaiseEvent Click
End Sub

Private Sub lblCap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblCap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    With UserControl
        If IsDft = False Then
            .BackColor = &HF5F5F5
            shpBrd.BorderColor = &HC0C0C0
            lblCap.ForeColor = &H808080
        Else
            .BackColor = &HFA8C5A
            shpBrd.BorderColor = &HD27800
            lblCap.ForeColor = &HF5F5F5
        End If
    End With
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.BackColor = shpBrd.BorderColor
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With UserControl
        If IsDft = False Then
            .BackColor = &HF5F5F5
        Else
            .BackColor = &HFA8C5A
        End If
    End With
End Sub

Private Sub UserControl_Resize()
    shpBrd.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    lblCap.Move 0, (UserControl.ScaleHeight - lblCap.Height) / 2, UserControl.ScaleWidth
End Sub

Public Property Get Caption() As String
    Caption = lblCap.Caption
End Property

Public Property Let Caption(ByVal nCap As String)
    PropertyChanged "Caption"
    lblCap.Caption = nCap
End Property

Public Property Get Default() As Boolean
    Default = IsDft
End Property

Public Property Let Default(ByVal nDft As Boolean)
    PropertyChanged "Default"
    IsDft = nDft
    UserControl_Initialize
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblCap.Caption = PropBag.ReadProperty("Caption", "Button")
    IsDft = PropBag.ReadProperty("Default", False)
    UserControl_Initialize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", lblCap.Caption
    PropBag.WriteProperty "Default", IsDft
End Sub

