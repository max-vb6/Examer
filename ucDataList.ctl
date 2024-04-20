VERSION 5.00
Begin VB.UserControl ucDataList 
   BackColor       =   &H00F5F5F5&
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3615
   ScaleWidth      =   5895
   Begin VB.Timer tmrScr 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin ExamerSvr.ucDataItem diChoice 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Caption         =   "选择题得分 XXX/XXX 分"
      BackColor       =   14737632
      TipTitle        =   "状态: 未全部回答"
      TipText         =   ""
      TipQue          =   "Question"
      QueVisible      =   0   'False
   End
   Begin VB.PictureBox picScr 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   480
      Width           =   135
      Begin VB.Label lblScr 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2655
      ScaleWidth      =   5775
      TabIndex        =   1
      Top             =   480
      Width           =   5775
      Begin ExamerSvr.ucDataItem diText 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   873
         Caption         =   "简答题 XXX"
         BackColor       =   16119285
         TipTitle        =   "用户回答 (未评分)"
         TipText         =   ""
         TipQue          =   "Question"
         QueVisible      =   -1  'True
      End
   End
End
Attribute VB_Name = "ucDataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event MarkSet()

Sub ResizeList()
    On Error Resume Next
    Dim i As Long, lNow As Long
    If diText.Count <= 1 Then Exit Sub
    picList.Move picScr.Width, diChoice.Height, UserControl.ScaleWidth - picScr.Width, 495 * (diText.Count - 1)
    lNow = 0
    For i = 1 To diText.UBound
        diText(i).Move 0, lNow * 495, picList.ScaleWidth, 495
        lNow = lNow + 1
    Next i
    With lblScr
        If picList.Height > UserControl.ScaleHeight - diChoice.Height Then
            .Move 0, 0, lblScr.Width, lblScr.Height / 3
            .Tag = picList.Height - (UserControl.ScaleHeight - diChoice.Height)
            .Visible = True
        Else
            .Visible = False
            .Tag = ""
        End If
    End With
End Sub

Function GetMarked(sIP As String) As Long
    On Error GoTo GMError
    
    If picList.Tag = "" Then GoTo GMError        'Tag 内存有 IP 信息
    Dim i As Long, lPlus As Long
    lPlus = 0
    i = GetUserQueNum(sIP)
    If i <> CLng(ReadLib("QueNum")) Then
        GoTo GMError
    Else
        For i = 1 To CLng(ReadLib("QueNum"))
            If Not IsNumeric(ReadUser(sIP, i - 1)) Then
                GoTo GMError
            Else
                lPlus = lPlus + CLng(ReadUser(sIP, i - 1))
            End If
        Next i
    End If
    
    GetMarked = lPlus
    
    Exit Function
GMError:
    GetMarked = -1
End Function

Function ReloadData(sIP As String) As Boolean
    On Error GoTo RDError
    
    Dim i As Long
    If diText.Count > 1 Then
        For i = 1 To diText.UBound
            Unload diText(i)
        Next i
    End If
    picList.Tag = ""
    i = GetUserQueNum(sIP)
    If i = 0 Then
        ReloadData = False
        Exit Function
    Else
        Dim lScrCh(2) As Long, sChc As String, lTxtNum As Long, sTmp As String
        lScrCh(0) = 0: lScrCh(1) = 0: lScrCh(2) = 0
        lTxtNum = 1
        picList.Tag = sIP
        For i = 0 To CLng(ReadLib("QueNum"))
            If LCase(Left(GetQueType(ReadLib("Que" & i), 5), 2)) = "ch" Then
                lScrCh(1) = lScrCh(1) + CLng(GetQueType(ReadLib("Que" & i), 6))
                lScrCh(2) = lScrCh(2) + 1
            End If
        Next i
        For i = 1 To GetUserQueNum(sIP)
            sTmp = GetQueType(ReadLib("Que" & i), 5)
            If LCase(Left(sTmp, 2)) = "ch" Then
                lScrCh(0) = lScrCh(0) + CLng(ReadUser(sIP, i - 1))
                sChc = sChc & "第 " & i & " 题: " & ReadUser(sIP, i - 1) & " 分" & vbCrLf
            ElseIf LCase(sTmp) = "text" Then
                Load diText(lTxtNum)
                With diText(lTxtNum)
                    .Tag = i
                    .Caption = "简答题第 " & i & " 题"
                    .TipQue = Replace(Replace(GetQueType(ReadLib("Que" & i), 0), vbCrLf, ""), "<br>", vbCrLf)
                    If IsNumeric(ReadUser(sIP, i - 1)) Then
                        .TipTitle = "用户答案 (本题 " & GetQueType(ReadLib("Que" & i), 6) & " 分，已评分)"
                        .SetMarkText ReadUser(sIP, i - 1)
                        .TipText = "[您可以修改评分]"
                    Else
                        .TipTitle = "用户答案 (本题 " & GetQueType(ReadLib("Que" & i), 6) & " 分，未评分)"
                        .SetMarkText ""
                        .TipText = ReadUser(sIP, i - 1)
                    End If
                    .Visible = True
                End With
                lTxtNum = lTxtNum + 1
            End If
        Next i
        sChc = Replace(sChc & " ", vbCrLf & " ", "")
        With diChoice
            .Caption = "选择题得分 " & CStr(lScrCh(0)) & "/" & CStr(lScrCh(1)) & " 分"
            If Len(sChc) - Len(Replace(sChc, "分", "")) < lScrCh(2) Then
                .TipTitle = "状态: 未全部回答"
            Else
                .TipTitle = "状态: 已全部回答"
            End If
            .TipText = sChc
        End With
        ResizeList
    End If
    
    ReloadData = True
    
    Exit Function
RDError:
    ReloadData = False
End Function

Private Sub diText_MarkSet(Index As Integer, lMark As Long)
    If picList.Tag = "" Then Exit Sub    'Tag 内存有 IP 信息
    If lMark > CLng(GetQueType(ReadLib("Que" & diText(Index).Tag), 6)) Then Beep: Exit Sub
    WriteUser picList.Tag, CLng(diText(Index).Tag) - 1, CStr(lMark)
    If Index < diText.UBound Then diText(Index + 1).SetTextFocus
    RaiseEvent MarkSet
End Sub

Private Sub lblScr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Static oy!
    With lblScr
        If (Not IsNumeric(.Tag)) Or tmrScr.Enabled Then Exit Sub
        If Button = 1 Then
            .Top = .Top - oy + Y
        Else
            oy = Y
        End If
        picList.Top = diChoice.Height - CSng(.Tag) * (.Top / (picScr.ScaleHeight - .Height))
    End With
End Sub

Private Sub lblScr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblScr.Top < 0 Or lblScr.Top > picScr.ScaleHeight - lblScr.Height Then
        tmrScr.Enabled = True
    End If
End Sub

Private Sub tmrScr_Timer()
    With lblScr
        If .Top < 0 Then
            .Top = .Top + GetMoveNum(0, .Top, 5)
            If GetMoveNum(0, .Top, 5) = 0 Then
                picList.Top = diChoice.Height
                tmrScr.Enabled = False
            End If
        ElseIf .Top > picScr.ScaleHeight - .Height Then
            .Top = .Top + GetMoveNum(picScr.ScaleHeight - .Height, .Top, 5)
            If GetMoveNum(picScr.ScaleHeight - .Height, .Top, 5) = 0 Then
                picList.Top = diChoice.Height - CSng(.Tag)
                tmrScr.Enabled = False
            End If
        End If
        picList.Top = diChoice.Height - CSng(.Tag) * (.Top / (picScr.ScaleHeight - .Height))
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With UserControl
        diChoice.Move 0, 0, .ScaleWidth, 495
        picScr.Move 0, diChoice.Height, 135, .ScaleHeight - diChoice.Height
        lblScr.Width = picScr.ScaleWidth
        ResizeList
    End With
End Sub
