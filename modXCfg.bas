Attribute VB_Name = "modXCfg"
Option Explicit
'==============================
'       XCfg ´æÈ¡º¯ÊýÄ£¿é
'         By MaxXSoft
'==============================
Const lVersion As Long = 0

Public Function MyPath() As String
    Dim sPath As String
    sPath = App.Path
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    MyPath = sPath
End Function

Function ReadCon(Item As String) As String
    ReadCon = LoadItem(Item, MyPath & "settings.xcfg")
End Function

Function SaveCon(Item As String, Value As String) As Long
    SaveCon = SaveItem(Item, Value, MyPath & "settings.xcfg")
End Function

Function ReadLib(Item As String) As String
    ReadLib = LoadItem(Item, MyPath & ReadCon("Qlib"))
End Function

Function SaveLib(Item As String, Value As String) As Long
    SaveLib = SaveItem(Item, Value, MyPath & ReadCon("Qlib"))
End Function

Public Function LoadItem(Name As String, Path As String) As String
    On Error GoTo LoadItemError:
    
    If Path = "" Or Dir(Path) = "" Then GoTo LoadItemError
    Dim Pb As New PropertyBag
    ReadPb Pb, Path
    If Pb.ReadProperty("/version") = "" Or CLng(Pb.ReadProperty("/version")) > lVersion Then
        GoTo LoadItemError
    Else
        LoadItem = Pb.ReadProperty(Name)
    End If
    Set Pb = Nothing
    
    Exit Function
LoadItemError:
    LoadItem = ""
End Function

Public Function SaveItem(Name As String, Value As String, Path As String) As Long
    On Error GoTo SaveItemError
    
    Dim Pb As New PropertyBag, SavePb As New PropertyBag
    ReadPb Pb, Path
    If Pb.ReadProperty("/version") = "" Or CLng(Pb.ReadProperty("/version")) > lVersion Then
        GoTo SaveItemError
    End If
    SavePb.WriteProperty "/version", lVersion
    
    Dim sIndexs() As String, i As Long
    sIndexs = Split(Pb.ReadProperty("/index"), "/")
    For i = 0 To UBound(sIndexs)
        If sIndexs(i) = "" Then Exit For
        If sIndexs(i) = Name Then
            SavePb.WriteProperty Name, Value
        Else
            SavePb.WriteProperty sIndexs(i), Pb.ReadProperty(sIndexs(i))
        End If
    Next i
    SavePb.WriteProperty "/index", Pb.ReadProperty("/index")
    
    Dim lFreeNum As Long, bytData() As Byte
    lFreeNum = FreeFile
    If Dir(Path) <> "" Then Kill Path
    Open Path For Binary As lFreeNum
        bytData = SavePb.Contents
        Put lFreeNum, 1, bytData
    Close lFreeNum
    
    Set Pb = Nothing
    Set SavePb = Nothing
    
    Exit Function
SaveItemError:
    SaveItem = IIf(Err.Number <> 0, Err.Number, -1)
End Function

Private Function ReadPb(PrBag As PropertyBag, Path As String) As Long
    On Error GoTo ReadPbError
    
    Dim lFreeNum As Long, bytData() As Byte
    lFreeNum = FreeFile
    Open Path For Binary As lFreeNum
        If LOF(lFreeNum) > 0 Then
            ReDim bytData(LOF(lFreeNum) - 1)
            Get lFreeNum, 1, bytData
            PrBag.Contents = bytData
        End If
    Close lFreeNum
    
    Exit Function
ReadPbError:
    ReadPb = Err.Number
End Function
