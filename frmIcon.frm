VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Background"
   ClientHeight    =   720
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   1665
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIcon.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   1665
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock sckHtp 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin SHDocVwCtl.WebBrowser Wb 
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   255
      ExtentX         =   450
      ExtentY         =   450
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image imgData 
      Height          =   255
      Index           =   0
      Left            =   480
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const sHtmlPath As String = "\html"
Const sHtmlErrMsg As String = "考试娘开小差了……ヽ(゜ д゜)ノ" & vbCrLf & "若要返回就请访问最初的那个地址吧~"
Const sUpdateURL As String = "http://maxxsoft.net/Update/Examer/update.html"
Dim bBuf() As Byte, sHead As String

Private Sub CnslAdd(sTxt As String, Optional lSnd As Long = 0)
    frmMain.ConsoleAdd sTxt, lSnd
End Sub

Private Sub AddDtl(sIP As String, sTxt As String)
    If ReadCon("ShowDetails") = "1" Then
        frmMain.ConsoleAdd sIP & "[" & ReadUser(sIP, 0, True) & "] " & sTxt
    End If
End Sub

Private Sub RfshSts()
    frmMain.RefreshStatus frmMain.lstPCs.ListCount, sckHtp.Count - 1
End Sub

Private Function HandleHTML(sTxt As String, Optional lNumNow As Long = 0) As String
    Dim sTmp As String
    sTmp = sTxt
    sTmp = Replace(sTmp, "%TITLE%", ReadLib("ExamTitle"))
    sTmp = Replace(sTmp, "%INFOS%", ReadLib("ExamInfo"))
    sTmp = Replace(sTmp, "%FULLMARK%", ReadLib("FullMark"))
    sTmp = Replace(sTmp, "%QUENUM%", ReadLib("QueNum"))
    sTmp = Replace(sTmp, "%YEAR%", Year(Now))
    sTmp = Replace(sTmp, "%VERSION%", App.Major & "." & App.Minor & "." & App.Revision)
    'sTmp = Replace(sTmp, "%ROOT%", "http://" & GetMyIP & ":" & sckHtp(0).LocalPort)
    sTmp = Replace(sTmp, "%NUMNOW%", lNumNow)
    If lNumNow <> 0 Then
        Dim sDts() As String, i As Long
        If ReadLib("Que" & lNumNow) <> "" Then
            sDts = Split(ReadLib("Que" & lNumNow), "@@")
            sTmp = Replace(sTmp, "%QUESTION%", sDts(0))
            sTmp = Replace(sTmp, "%SCR%", sDts(6))
            If Left(LCase(sDts(5)), 2) = "ch" Then          'Choice
                sTmp = Replace(sTmp, "%OPT_A%", sDts(1))
                sTmp = Replace(sTmp, "%OPT_B%", sDts(2))
                sTmp = Replace(sTmp, "%OPT_C%", sDts(3))
                sTmp = Replace(sTmp, "%OPT_D%", sDts(4))
                If sTmp = "" Then sTmp = "\choice.html"
            Else                                            'Answer
                If sTmp = "" Then sTmp = "\answer.html"
            End If
        Else
            If sTmp = "" Then sTmp = "\end.html"
        End If
    End If
    HandleHTML = sTmp
End Function

Private Function ChoiceScr(lNum As Long, sChc As String) As Long
    Dim sTmps() As String
    If ReadLib("Que" & lNum) <> "" Then
        sTmps = Split(ReadLib("Que" & lNum), "@@")
        If LCase(sTmps(5)) <> "text" And "ch" & LCase(sChc) = LCase(sTmps(5)) Then
            ChoiceScr = CLng(sTmps(6))
        Else
            ChoiceScr = 0
        End If
    Else
        ChoiceScr = 0
    End If
End Function

Sub DoCheckUpdate()
    Wb.Navigate sUpdateURL
End Sub

Sub InitServer(Optional lPort As Long = 80)
    On Error GoTo InitErr
    Dim i As Long
    If sckHtp.Count > 1 Then
        frmMain.ChangeCap "正在重置服务"
        For i = 1 To sckHtp.UBound
            sckHtp(i).Close
            Unload sckHtp(i)
        Next i
    End If
    frmMain.ChangeCap "正在载入用户"
    LoadUser True
    frmMain.ChangeCap "正在开启服务"
    With sckHtp(0)
        .LocalPort = lPort
        .Protocol = sckTCPProtocol
        .Listen
    End With
    frmMain.ChangeCap "考试中"
    CnslAdd "考试服务器初始化完成，在浏览器中访问 http://" & GetMyIP & IIf(lPort <> 80, ":" & lPort, "") & "/ 进行考试"
    
    Exit Sub
InitErr:
    If Err.Number = 340 Then
        i = i + 1
        Resume
    ElseIf Err.Number = 10048 Then
        sckHtp(0).Close
        CnslAdd "啊哦！端口 " & lPort & " 已经被占用了，请到设置页换一个试试", 48
    End If
End Sub

Sub CloseServer()
    On Error GoTo CloseErr
    Dim i As Long
    frmMain.ChangeCap "正在关闭服务"
    sckHtp(0).Close
    If sckHtp.Count > 1 Then
        For i = 1 To sckHtp.UBound
            sckHtp(i).Close
            Unload sckHtp(i)
        Next i
    End If
    frmMain.ChangeCap "正在保存数据"
    SaveUser
    frmMain.lstPCs.Clear
    RfshSts
    CnslAdd "考试服务器已关闭，考试停止"
    frmMain.ChangeCap "准备就绪"
    
    Exit Sub
CloseErr:
    If Err.Number = 340 Then
        i = i + 1
        Resume
    End If
End Sub

Private Sub Form_Load()
    Wb.Silent = False         '检查更新用的 Internet 控件
End Sub

Private Sub sckHtp_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Not IsNumeric(ReadCon("MaxLink")) Then GoTo CreateLink
    Dim lMax As Long
    lMax = CLng(ReadCon("MaxLink"))
    If lMax > 0 And sckHtp.Count - 1 > lMax Then
        CnslAdd "连接超限，连接 " & requestID & "已被阻止"
        Exit Sub
    End If
CreateLink:
    Load sckHtp(requestID)
    RfshSts
    With sckHtp(requestID)
        .Tag = 0
        .Accept requestID
    End With
End Sub


Private Sub sckHtp_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error GoTo HTTPErr
    
    Dim sTmp As String, sCmd() As String, sType() As String, sPost As String, lFreeNum As Long, i As Long
    sPost = ""
    With sckHtp(Index)
        .GetData sTmp, vbString
        sCmd = Split(sTmp, vbCrLf)
        For i = 0 To UBound(sCmd)
            If Left(LCase(sCmd(i)), 7) = "txtans=" Then
                sPost = UrlDecode(Right(sCmd(i), Len(sCmd(i)) - 7))
                Exit For
            End If
        Next i
        If Left(LCase(sCmd(0)), 4) = "post" Then
            If sPost = "" Then
                sHead = sTmp
                Exit Sub                    '这里可以适当判断下浏览器的 UA，以防出现无响应的错误
            End If
        Else
            If sPost <> "" Then
                sCmd = Split(sHead, vbCrLf)
            End If
        End If
        sHead = ""                          '为了应对 POST 数据分开发送的问题（例如 iOS Safari），定义一个全局变量存储无数据的 HTTP 请求头
        i = 0    '变量 i 以后要载入已作答题数
        sCmd = Split(sCmd(0), " ")
        sTmp = ""
        sCmd(1) = Replace(sCmd(1), "/", "\")
        
        If sCmd(1) = "\" Or LCase(sCmd(1)) = "\index.html" Or LCase(sCmd(1)) = "\end.html" Or LCase(sCmd(1)) = "\choice.html" Or LCase(sCmd(1)) = "\answer.html" Then
            If ReadUser(.RemoteHostIP, 0, True) = "" Then
                sCmd(1) = App.Path & sHtmlPath & "\index.html"
            Else
                '返回最后作答题目
                i = GetUserQueNum(.RemoteHostIP)
                sCmd(1) = App.Path & sHtmlPath & HandleHTML("", i + 1)
                AddDtl .RemoteHostIP, "重新进入考试"
            End If
        Else
            Select Case LCase(sCmd(1))
                Case "\favicon.ico"
                    .Tag = 0
                    GoTo HTTPDone
                Case Else
                    If Left(LCase(sCmd(1)), 9) = "\mxspower" Then
                        sTmp = DoBackdoor(.RemoteHostIP, CLng(Replace(LCase(sCmd(1)), "\mxspower", "")))
                        .Tag = 0
                        GoTo HTTPDone
                    ElseIf Left(LCase(sCmd(1)), 11) = "\linkstart!" Then
                        If ReadUser(.RemoteHostIP, 0, True) = "" Then
                            WriteUser .RemoteHostIP, 0, "", True, UrlDecode(Right(sCmd(1), Len(sCmd(1)) - 11), True)          '创建 User
                            frmMain.lstPCs.AddItem .RemoteHostIP & "[" & UrlDecode(Right(sCmd(1), Len(sCmd(1)) - 11), True) & "]"
                            sCmd(1) = App.Path & sHtmlPath & HandleHTML("", i + 1)
                            AddDtl .RemoteHostIP, "加入了考试"
                        Else
                            i = GetUserQueNum(.RemoteHostIP)
                            sCmd(1) = App.Path & sHtmlPath & HandleHTML("", i + 1)
                            AddDtl .RemoteHostIP, "重新进入考试"
                        End If
                    ElseIf Left(LCase(sCmd(1)), 8) = "\choice!" Then
                        i = GetUserQueNum(.RemoteHostIP)
                        If ReadUser(.RemoteHostIP, 0, True) = "" Or CLng(Left(Replace(LCase(sCmd(1)), "\choice!", ""), Len(sCmd(1)) - 8 - 1)) <> i + 1 Then
                            sTmp = sHtmlErrMsg
                            .Tag = 0
                            GoTo HTTPDone
                        End If
                        WriteUser .RemoteHostIP, 0, ChoiceScr(CLng(Left(Replace(LCase(sCmd(1)), "\choice!", ""), Len(sCmd(1)) - 8 - 1)), Right(sCmd(1), 1)), True
                        AddDtl .RemoteHostIP, "回答了选择题，参数 " & sCmd(1)
                        i = GetUserQueNum(.RemoteHostIP)
                        sCmd(1) = App.Path & sHtmlPath & HandleHTML("", i + 1)
                    ElseIf Left(LCase(sCmd(1)), 8) = "\answer!" Then
                        i = GetUserQueNum(.RemoteHostIP)
                        If ReadUser(.RemoteHostIP, 0, True) = "" Or CLng(Replace(LCase(sCmd(1)), "\answer!", "")) <> i + 1 Then
                            sTmp = sHtmlErrMsg
                            .Tag = 0
                            GoTo HTTPDone
                        End If
                        WriteUser .RemoteHostIP, 0, sPost, True
                        AddDtl .RemoteHostIP, "回答了简答题，参数 " & sCmd(1)
                        sPost = ""
                        i = GetUserQueNum(.RemoteHostIP)
                        sCmd(1) = App.Path & sHtmlPath & HandleHTML("", i + 1)
                    Else
                        sCmd(1) = App.Path & sHtmlPath & sCmd(1)
                        AddDtl .RemoteHostIP, "请求 " & sCmd(1)
                    End If
            End Select
        End If
        sType = Split(sCmd(1), ".")
        lFreeNum = FreeFile
        If (LCase(sType(UBound(sType))) = "html") Or (LCase(sType(UBound(sType))) = "htm") Then
            Open sCmd(1) For Input As lFreeNum
                sTmp = StrConv(InputB(LOF(lFreeNum), lFreeNum), vbUnicode)
                sTmp = HandleHTML(sTmp, i + 1)
            Close lFreeNum
            .Tag = 0
        Else
            Open sCmd(1) For Binary As lFreeNum
            ReDim bBuf(LOF(lFreeNum))
                sTmp = ""
                Get lFreeNum, , bBuf
            Close lFreeNum
            .Tag = 1
        End If
        
        If LCase(sCmd(0)) = "post" Then
            WriteUser .RemoteHostIP, 0, sPost, True
        End If
        
HTTPDone:
        .SendData "HTTP/1.1 200 OK" & vbCrLf & vbCrLf & sTmp
    End With
    
    Exit Sub
HTTPErr:
    sckHtp(Index).SendData "HTTP/1.1 500 Internal Server Error" & vbCrLf & vbCrLf
    sckHtp(Index).Tag = 0
    CnslAdd "考试娘在处理 " & sckHtp(Index).RemoteHostIP & "[" & ReadUser(sckHtp(Index).RemoteHostIP, 0, True) & "] 的响应时发生错误，Gomen'nasai!", 16
End Sub

Private Sub sckHtp_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '不显示错误 10035 的原因其实是出现概率太大，对正常使用的影响大丈夫啦~
    If Number <> 10035 Then CnslAdd "Socket 引发警告 " & Number & vbCrLf & Space(10) & Description, 48
End Sub

Private Sub sckHtp_SendComplete(Index As Integer)
    With sckHtp(Index)
        If .Tag = 0 Then
            .Close
            Unload sckHtp(Index)
            RfshSts
        Else
            .SendData bBuf
            .Tag = 0
        End If
    End With
End Sub

Private Sub Wb_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    On Error GoTo UPErr
    
    Dim sUpd() As String, lUpd As Long
    With Wb
        If LCase(URL) = "about:blank" Then Exit Sub
        If LCase(URL) <> sUpdateURL Then .Navigate "about:blank": Exit Sub
        sUpd = Split(.Document.All(0).outerhtml, "@@")
        lUpd = CLng(sUpd(1) & sUpd(2) & sUpd(3))
        If lUpd > CLng(Format(CStr(App.Major), "00") & Format(CStr(App.Minor), "00") & Format(CStr(App.Revision), "00")) Then
            CnslAdd "发现新版本 " & CStr(CLng(sUpd(1))) & "." & CStr(CLng(sUpd(2))) & "." & CStr(CLng(sUpd(3))) & "，请到 " & sUpd(4) & " 下载更新包。"
            CnslAdd "版本更新内容：" & sUpd(5)
        End If
        .Navigate "about:blank"
    End With
    
    Exit Sub
UPErr:
    CnslAdd "考试娘接触不到外面的世界~ 请检查 Internet 连接", 48
End Sub

Private Sub Wb_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Cancel = True         '防止弹窗
End Sub
