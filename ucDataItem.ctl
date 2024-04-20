VERSION 5.00
Begin VB.UserControl ucDataItem 
   Appearance      =   0  'Flat
   BackColor       =   &H00D7D7D7&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   615
   ScaleWidth      =   4575
   ToolboxBitmap   =   "ucDataItem.ctx":0000
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      BackColor       =   &H00D7D7D7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2880
      ScaleHeight     =   615
      ScaleWidth      =   1335
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      Begin VB.TextBox txtMark 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Text            =   "0"
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "评分"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.Label lblDtl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[...]"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D27800&
      Height          =   360
      Left            =   4200
      TabIndex        =   0
      Top             =   60
      Width           =   285
   End
End
Attribute VB_Name = "ucDataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event MarkSet(lMark As Long)

Dim sTipTitle As String, sTipText As String, sTipQue As String
Dim bShow As Boolean

Sub SetMarkText(sMark As String)
    If Not IsNumeric(sMark) Then Exit Sub
    txtMark.Text = sMark
End Sub

Sub SetTextFocus()
    If picText.Visible Then
        txtMark.SetFocus
        txtMark.SelStart = Len(txtMark.Text)
    End If
End Sub

Private Sub lblDtl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bShow Then
        With Tip
            .Title = "题目内容"
            .TipText = sTipQue
            .Create UserControl.hwnd
        End With
        bShow = True
    End If
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bShow Then
        With Tip
            .Title = sTipTitle
            .TipText = sTipText
            .Create UserControl.hwnd
        End With
        bShow = True
    End If
End Sub

Private Sub txtMark_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
            If IsNumeric(txtMark.Text) Then
                RaiseEvent MarkSet(CLng(txtMark.Text))
            Else
                txtMark.Text = ""
                Beep
            End If
        Case 48 To 57
            Exit Sub
        Case vbKeyBack, vbKeyDelete
            Exit Sub
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub UserControl_Initialize()
    picText.BackColor = UserControl.BackColor
    picText.Visible = lblDtl.Visible
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bShow Then
        Tip.Destroy
        bShow = False
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With UserControl
        lblCap.Move 120, (.ScaleHeight - lblCap.Height) / 2
        lblDtl.Move lblCap.Left + lblCap.Width + 120, lblCap.Top
        lblInfo.Move .ScaleWidth - lblInfo.Width - 120, lblCap.Top - 60
        picText.Move .ScaleWidth - lblInfo.Width - picText.Width - 120, 0, lblShow.Width + txtMark.Width + 360, .ScaleHeight
        lblShow.Move 120, (picText.ScaleHeight - lblShow.Height) / 2
        txtMark.Move lblShow.Left + lblShow.Width + 120, (picText.ScaleHeight - txtMark.Height) / 2
    End With
End Sub

Public Property Get Caption() As String
    Caption = lblCap.Caption
End Property

Public Property Let Caption(ByVal nCap As String)
    PropertyChanged "Caption"
    lblCap.Caption = nCap
    UserControl_Resize
End Property

Public Property Get TipTitle() As String
    TipTitle = sTipTitle
End Property

Public Property Let TipTitle(ByVal sTlt As String)
    PropertyChanged "TipTitle"
    sTipTitle = sTlt
End Property

Public Property Get TipText() As String
    TipText = sTipText
End Property

Public Property Let TipText(ByVal sTxt As String)
    PropertyChanged "TipText"
    sTipText = sTxt
End Property

Public Property Get TipQue() As String
    TipQue = sTipQue
End Property

Public Property Let TipQue(ByVal sQue As String)
    PropertyChanged "TipQue"
    sTipQue = sQue
End Property

Public Property Get QueVisible() As Boolean
    QueVisible = lblDtl.Visible
End Property

Public Property Let QueVisible(ByVal bVs As Boolean)
    PropertyChanged "QueVisible"
    lblDtl.Visible = bVs
    picText.Visible = bVs
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal cClr As OLE_COLOR)
    PropertyChanged "BackColor"
    UserControl.BackColor = cClr
    picText.BackColor = cClr
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        lblCap.Caption = .ReadProperty("Caption", "Text")
        UserControl.BackColor = .ReadProperty("BackColor")
        sTipTitle = .ReadProperty("TipTitle", "Detail")
        sTipText = .ReadProperty("TipText", "Text")
        sTipQue = .ReadProperty("TipQue", "Question")
        lblDtl.Visible = .ReadProperty("QueVisible", False)
    End With
    UserControl_Initialize
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", lblCap.Caption
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "TipTitle", sTipTitle
        .WriteProperty "TipText", sTipText
        .WriteProperty "TipQue", sTipQue
        .WriteProperty "QueVisible", lblDtl.Visible
    End With
End Sub
