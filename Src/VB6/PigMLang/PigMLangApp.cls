VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PigMLangApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MLSetItems"
'*******************************************************
'* Name: MultiLangApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Piggy multilingualism applications
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Created: 11/5/2020
'*******************************************************
Option Explicit

Const ML_EXT_NAME As String = ".ml" '多语言包文件的扩展名   Extension of multi language package file
Private mvarCurrentLangID As Long '当前的语言代码   Current language code
Private mvarLangIDList As String '本语言字典下可用的语言ID列表  List of language IDs available in this language dictionary
Private mvarBindFormList As String '已绑定的窗体列表    List of bound forms
Private mvarMLSetItems As MLSetItems '多语言文本集  Multilingual text set
Private mstrDictFilePath As String '多语言包的路径  Path of multi language package
Private mstrDictFileTitle As String '多语言包的前缀 Prefixes for multilingual packages

'Log processing-Begin
Private moPigLog As PigLog

Public Sub InitLog(ByVal LogDir As String, ByVal LogFileTitle As String)
    Set moPigLog = New PigLog
    moPigLog.Init "pFileSystemObject", LogDir, LogFileTitle
End Sub

Public Property Get LastErr() As String
    LastErr = moPigLog.LastErr
End Property

Public Property Get LogDir() As String
    LogDir = moPigLog.LogDir
End Property

Public Function GetFuncDemo(ByVal FuncName As String, ByVal FuncPara As String) As String
    GetFuncDemo = moPigLog.GetFuncDemo(FuncName, FuncPara)
End Function

Public Function GetSubDemo(ByVal SubName As String, ByVal SubPara As String) As String
    GetSubDemo = moPigLog.GetSubDemo(SubName, SubPara)
End Function

'Log processing-End

Public Sub ClearBindForm()
    mvarBindFormList = ""
End Sub

Public Property Get MLSetItems() As MLSetItems
    Set MLSetItems = mvarMLSetItems
End Property


Friend Property Set MLSetItems(vData As MLSetItems)
    Set mvarMLSetItems = vData
End Property

Private Sub Class_Initialize()
    Me.InitLog "", ""
    mvarCurrentLangID = GetUserDefaultLangID()
    InitMLSetItems
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mvarMLSetItems = Nothing
    On Error GoTo 0
End Sub


Private Sub InitMLSetItems()
    On Error GoTo ErrOcc:
    Set mvarMLSetItems = New MLSetItems
    With mvarMLSetItems
        .Add CLng("&H0000"), Me
        .Add CLng("&H0400"), Me
        .Add CLng("&H0401"), Me
        .Add CLng("&H0801"), Me
        .Add CLng("&H0c01"), Me
        .Add CLng("&H1001"), Me
        .Add CLng("&H1401"), Me
        .Add CLng("&H1801"), Me
        .Add CLng("&H1c01"), Me
        .Add CLng("&H2001"), Me
        .Add CLng("&H2401"), Me
        .Add CLng("&H2801"), Me
        .Add CLng("&H2c01"), Me
        .Add CLng("&H3001"), Me
        .Add CLng("&H3401"), Me
        .Add CLng("&H3801"), Me
        .Add CLng("&H3c01"), Me
        .Add CLng("&H4001"), Me
        .Add CLng("&H0402"), Me
        .Add CLng("&H0403"), Me
        .Add CLng("&H0404"), Me
        .Add CLng("&H0804"), Me
        .Add CLng("&H0c04"), Me
        .Add CLng("&H1004"), Me
        .Add CLng("&H0405"), Me
        .Add CLng("&H0406"), Me
        .Add CLng("&H0407"), Me
        .Add CLng("&H0807"), Me
        .Add CLng("&H0c07"), Me
        .Add CLng("&H1007"), Me
        .Add CLng("&H1407"), Me
        .Add CLng("&H0408"), Me
        .Add CLng("&H0409"), Me
        .Add CLng("&H0809"), Me
        .Add CLng("&H0c09"), Me
        .Add CLng("&H1009"), Me
        .Add CLng("&H1409"), Me
        .Add CLng("&H1809"), Me
        .Add CLng("&H1c09"), Me
        .Add CLng("&H2009"), Me
        .Add CLng("&H2409"), Me
        .Add CLng("&H2809"), Me
        .Add CLng("&H2c09"), Me
        .Add CLng("&H040a"), Me
        .Add CLng("&H080a"), Me
        .Add CLng("&H0c0a"), Me
        .Add CLng("&H100a"), Me
        .Add CLng("&H140a"), Me
        .Add CLng("&H180a"), Me
        .Add CLng("&H1c0a"), Me
        .Add CLng("&H200a"), Me
        .Add CLng("&H240a"), Me
        .Add CLng("&H280a"), Me
        .Add CLng("&H2c0a"), Me
        .Add CLng("&H300a"), Me
        .Add CLng("&H340a"), Me
        .Add CLng("&H380a"), Me
        .Add CLng("&H3c0a"), Me
        .Add CLng("&H400a"), Me
        .Add CLng("&H440a"), Me
        .Add CLng("&H480a"), Me
        .Add CLng("&H4c0a"), Me
        .Add CLng("&H500a"), Me
        .Add CLng("&H040b"), Me
        .Add CLng("&H040c"), Me
        .Add CLng("&H080c"), Me
        .Add CLng("&H0c0c"), Me
        .Add CLng("&H100c"), Me
        .Add CLng("&H140c"), Me
        .Add CLng("&H040d"), Me
        .Add CLng("&H040e"), Me
        .Add CLng("&H040f"), Me
        .Add CLng("&H0410"), Me
        .Add CLng("&H0810"), Me
        .Add CLng("&H0411"), Me
        .Add CLng("&H0412"), Me
        .Add CLng("&H0812"), Me
        .Add CLng("&H0413"), Me
        .Add CLng("&H0813"), Me
        .Add CLng("&H0414"), Me
        .Add CLng("&H0814"), Me
        .Add CLng("&H0415"), Me
        .Add CLng("&H0416"), Me
        .Add CLng("&H0816"), Me
        .Add CLng("&H0418"), Me
        .Add CLng("&H0419"), Me
        .Add CLng("&H041a"), Me
        .Add CLng("&H081a"), Me
        .Add CLng("&H0c1a"), Me
        .Add CLng("&H041b"), Me
        .Add CLng("&H041c"), Me
        .Add CLng("&H041d"), Me
        .Add CLng("&H081d"), Me
        .Add CLng("&H041e"), Me
        .Add CLng("&H041f"), Me
        .Add CLng("&H0421"), Me
        .Add CLng("&H0422"), Me
        .Add CLng("&H0423"), Me
        .Add CLng("&H0424"), Me
        .Add CLng("&H0425"), Me
        .Add CLng("&H0426"), Me
        .Add CLng("&H0427"), Me
        .Add CLng("&H0429"), Me
        .Add CLng("&H042a"), Me
        .Add CLng("&H042d"), Me
        .Add CLng("&H0436"), Me
        .Add CLng("&H0438"), Me
    End With
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Sub
ErrOcc:
    moPigLog.SetSubErrInf "InitMLSetItems", "", Err
    On Error GoTo 0
End Sub

Public Function MKDictString(frmAny As Object, Optional IsFont As Boolean = True) As String
'产生全部窗体的字典结构
'IsFont 是否更新控件的字体
Dim ctlAny As Control, i As Long
Dim strData As String, strTypeName As String, strTmp As String

    On Error GoTo ErrOcc:
    MKDictString = ""
    
    On Error Resume Next
    MKDictString = MKDictString & "{" & frmAny.Name & "}" & vbCrLf
    MKDictString = MKDictString & "[Me_Caption]=" & frmAny.Caption & vbCrLf
    MKDictString = MKDictString & "[Me_FontName]=" & frmAny.Font.Name & vbCrLf
    MKDictString = MKDictString & "[Me_FontSize]=" & frmAny.Font.Size & vbCrLf
    
    For Each ctlAny In frmAny.Controls
        strData = ""
        '控件类型名
        strTypeName = TypeName(ctlAny)
        Select Case strTypeName
        Case "CommonDialog", "Inet", "MAPIMessages", "MSComm", "MAPISession", "Timer", "Winsock", "SysInfo", "PictureClip", "ImageList"
            '没有界面，不处理
        Case Else
            '返回控件的属性文本
            strData = GetObjPropertyText(ctlAny, IsFont)
            '处理字控件集
            Select Case strTypeName
            Case "StatusBar"
                For i = 1 To ctlAny.Panels.Count
                    strTmp = "[" & ctlAny.Name & "_Panels(" & i & ")_ToolTipText]=" & ctlAny.Panels(i).ToolTipText & vbCrLf
                    strData = strData & strTmp
                Next
            Case "CoolBar"
                For i = 1 To ctlAny.Bands.Count
                    strTmp = "[" & ctlAny.Name & "_Bands(" & i & ")_Caption]=" & ctlAny.Bands(i).Caption & vbCrLf
                    strData = strData & strTmp
                Next
            Case "TabStrip"
                For i = 1 To ctlAny.Tabs.Count
                    strTmp = "[" & ctlAny.Name & "_Tabs(" & i & ")_Caption]=" & ctlAny.Tabs(i).Caption & vbCrLf
                    strData = strData & strTmp
                    strTmp = "[" & ctlAny.Name & "_Tabs(" & i & ")_ToolTipText]=" & ctlAny.Tabs(i).ToolTipText & vbCrLf
                    strData = strData & strTmp
                Next
            Case "Toolbar"
                For i = 1 To ctlAny.Buttons.Count
                    strTmp = "[" & ctlAny.Name & "_Buttons(" & i & ")_Caption]=" & ctlAny.Buttons(i).Caption & vbCrLf
                    strData = strData & strTmp
                    strTmp = "[" & ctlAny.Name & "_Buttons(" & i & ")_ToolTipText]=" & ctlAny.Buttons(i).ToolTipText & vbCrLf
                    strData = strData & strTmp
                Next
            End Select
            MKDictString = MKDictString & strData
        End Select
    Next
    moPigLog.ClearErr
    Set ctlAny = Nothing
    On Error GoTo 0
    Exit Function
ErrOcc:
    moPigLog.SetSubErrInf "MKDictString", "", Err
    Set ctlAny = Nothing
    On Error GoTo 0
End Function



Public Sub RefreshBindForm(frmAny As Object, Optional IsFont As Boolean = True)
'刷新绑定的窗体
'IsFont 是否更新控件的字体
Dim strFormName As String, ctlAny As Control, bolFind As Boolean
Dim oMLSetItem As MLSetItem, strKey As String, strValue As String
Dim strTypeName As String, strTmp As String, i As Long

    On Error GoTo ErrOcc:
    
    strFormName = "<" & frmAny.Name & ">"
    If InStr(mvarBindFormList, strFormName) = 0 Then
        Err.Raise -1, , "Form not Binded!"
    End If
    '这个窗体已绑定，检查信息集中有否登记
    bolFind = False
    For Each oMLSetItem In mvarMLSetItems
        strTmp = "<" & oMLSetItem.Key & ">"
        If "<" & oMLSetItem.Key & ">" = strFormName Then
            bolFind = True
            Exit For
        End If
    Next
    If bolFind = False Then
        Err.Raise -1, , " MultiLang Set not found(" & strFormName & ")"
    End If
    '修改窗体的属性
    On Error Resume Next
    '处理窗体的属性
    With frmAny
        '标题
        strKey = "Me_Caption"
        strValue = oMLSetItem.MLItems(strKey).Value
        If Err.Number <> 0 Then
            Err.Clear
        Else
            .Caption = strValue
        End If
        '字体
        strKey = "Me_FontName"
        strValue = oMLSetItem.MLItems(strKey).Value
        If Err.Number <> 0 Then
            Err.Clear
        Else
            .Font.Name = Trim(strValue)
        End If
        strKey = "Me_FontSize"
        strValue = oMLSetItem.MLItems(strKey).Value
        If Err.Number <> 0 Then
            Err.Clear
        Else
            If IsNumeric(strValue) Then
                .Font.Size = CCur(strValue)
            End If
        End If
    End With
    '修改控件的显示及字体
    For Each ctlAny In frmAny.Controls
        strTypeName = TypeName(ctlAny)
        Select Case strTypeName
        Case "CommonDialog", "Inet", "MAPIMessages", "MSComm", "MAPISession", "Timer", "Winsock", "SysInfo", "PictureClip", "ImageList"
            '没有界面，不处理
        Case Else
            SetObjPropertyText ctlAny, oMLSetItem, IsFont
        End Select
    Next

    moPigLog.ClearErr
    Set oMLSetItem = Nothing
    Set ctlAny = Nothing
    On Error GoTo 0
    Exit Sub
ErrOcc:
    moPigLog.SetSubErrInf "RefreshBindForm", "", Err
    Set oMLSetItem = Nothing
    Set ctlAny = Nothing
    On Error GoTo 0

End Sub


Public Property Get BindFormList() As String
    BindFormList = mvarBindFormList
End Property

Public Property Get LangIDList() As String
    LangIDList = mvarLangIDList
End Property


Public Property Get LangList() As String
Dim strTmp As String
    Do While True
        strTmp = gmGetStr(mvarLangIDList, "<", ">")
        If strTmp = "" Then Exit Do
        If IsNumeric(strTmp) Then
            strTmp = gmGetLangInf(CLng(strTmp))
            LangList = LangList & "<" & strTmp & ">"
        End If
    Loop
End Property

Public Property Get CurrentLangID() As Long
    CurrentLangID = mvarCurrentLangID
End Property




Public Sub BindForm(ByVal FormName As String)
'绑定一个窗体
    '检查是否已绑定
    FormName = "<" & FormName & ">"
    If InStr(mvarBindFormList, FormName) = 0 Then
        mvarBindFormList = mvarBindFormList & FormName
    End If
    
End Sub



Public Sub ChangLang(ByVal LangID As Long)
'切换语言系统，切换自动扫描目录，看看有没有新的语言库
'切换语言系统后要更新绑定的窗体的信息
    On Error GoTo ErrOcc:
    LoadDictFile mstrDictFilePath, mstrDictFileTitle, LangID
    If Me.LastErr <> "" Then Err.Raise -1, , Me.LastErr
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Sub
ErrOcc:
    moPigLog.SetSubErrInf "ChangLang", "", Err
    On Error GoTo 0
End Sub


Public Sub LoadDictFile(ByVal FilePath As String, ByVal FileTitle As String, Optional ByVal DefaultLangID As Long = 0)
'导入字典文件
'如果 DefaultLangID 为0则自动获取用户选择的语言标识
'如果没有 DefaultLangID 的语言标识，则自动选择可通用的
Dim oFolder As pFolder, oFile As pFile, strFileName As String, strLangID As String
Dim lngLangID As Long, strTmp As String, strTmp1 As String, strTmp2 As String, bolFind As Boolean
Dim oTextStream As pTextStream, strItem As String, oMLSetItem As MLSetItem
Dim bolIsMultiLine As Boolean, strKey As String, strValue As String
Dim strDictFilePath As String, lngLineNo As Long
Dim strSChinese As String, strTChinese As String, strEnglish As String
    
    On Error GoTo ErrOcc:
    
    '检查有多少个语言包
    If goFS Is Nothing Then Set goFS = New pFileSystemObject
10: Set oFolder = goFS.GetFolder(FilePath)
    mvarLangIDList = ""
    strTmp1 = LCase(FileTitle) & "_"
    mvarLangIDList = ""
    For Each oFile In oFolder.Files
20:     strFileName = LCase(oFile.Name)
        strLangID = gmGetStr(strFileName, strTmp1, ML_EXT_NAME, False)
        If IsNumeric(strLangID) Then
            lngLangID = CLng(strLangID)
            mvarLangIDList = mvarLangIDList & "<" & strLangID & ">"
        End If
    Next
    
    '确定当前语言代码
    If DefaultLangID = 0 Then
        '自动检查
        '先使用用户设置的语言
30:     DefaultLangID = GetUserDefaultLangID()
        strTmp = "<" & DefaultLangID & ">"
        If InStr(mvarLangIDList, strTmp) = 0 Then
            '如果没有则使用可互换的语言
            strTmp = gmGetAboutLangIDList(DefaultLangID)
            '检查是否有可互换的语言
            bolFind = False
            Do While True
                strLangID = gmGetStr(strTmp, "<", ">", True)
                If strLangID = "" Then Exit Do
                strTmp1 = "<" & strLangID & ">"
                If InStr(mvarLangIDList, strTmp1) <> 0 Then
                    DefaultLangID = CLng(strLangID)
                    bolFind = True
                    Exit Do
                End If
            Loop
            If bolFind = False Then
                DefaultLangID = GetUserDefaultLangID()
                strTmp = "<" & DefaultLangID & ">"
                strSChinese = gmGetAboutLangIDList(2052) '简体中文列表  Simplified Chinese list
                strTChinese = gmGetAboutLangIDList(1028) '繁体中文列表  Traditional Chinese list
                strEnglish = gmGetAboutLangIDList(9) '英文列表  English list
                If InStr(strSChinese, strTmp) <> 0 Then
                    '试一下换成繁体中文 Try traditional Chinese
                    strTmp1 = strTChinese
                    bolFind = False
                    Do While True
                        strLangID = gmGetStr(strTmp1, "<", ">", True)
                        If strLangID = "" Then Exit Do
                        strTmp2 = "<" & strLangID & ">"
                        If InStr(mvarLangIDList, strTmp2) <> 0 Then
                            DefaultLangID = CLng(strLangID)
                            bolFind = True
                            Exit Do
                        End If
                    Loop
                End If
                If bolFind = False Then
                    If InStr(strTChinese, strTmp) <> 0 Then
                        '试一下换成简体中文 Try to change it to simplified Chinese
                        strTmp1 = strSChinese
                        bolFind = False
                        Do While True
                            strLangID = gmGetStr(strTmp1, "<", ">", True)
                            If strLangID = "" Then Exit Do
                            strTmp2 = "<" & strLangID & ">"
                            If InStr(mvarLangIDList, strTmp2) <> 0 Then
                                DefaultLangID = CLng(strLangID)
                                bolFind = True
                                Exit Do
                            End If
                        Loop
                    End If
                End If
                If bolFind = False Then
                    '试一下换成英文 Try to change it to English
                    strTmp1 = strEnglish
                    bolFind = False
                    Do While True
                        strLangID = gmGetStr(strTmp1, "<", ">", True)
                        If strLangID = "" Then Exit Do
                        strTmp2 = "<" & strLangID & ">"
                        If InStr(mvarLangIDList, strTmp2) <> 0 Then
                            DefaultLangID = CLng(strLangID)
                            bolFind = True
                            Exit Do
                        End If
                    Loop
                End If
                If bolFind = False Then
                    '没办法了   No way out
                    Err.Raise -1, , "Can not find MultiFile!"
                End If
            End If
        End If
    Else
        '检查传入的语言ID是否有效   Check that the passed in language ID is valid
        strTmp = "<" & DefaultLangID & ">"
        If InStr(mvarLangIDList, strTmp) = 0 Then
            Err.Raise -1, , "MultiFile " & FilePath & FileTitle & "_" & DefaultLangID & ML_EXT_NAME & " not find!"
            GoTo ErrOcc:
        End If
    End If
    '导入语言包 Import language pack
    mstrDictFilePath = FilePath
    mstrDictFileTitle = FileTitle
    
    strDictFilePath = FilePath & "\" & FileTitle & "_"
    strDictFilePath = Replace(strDictFilePath, "\\", "\")
    mvarCurrentLangID = DefaultLangID
    strFileName = strDictFilePath & mvarCurrentLangID & ML_EXT_NAME
40: Set oTextStream = goFS.OpenTextFile(strFileName, pIOMode.ForAppending, True)
50: Set mvarMLSetItems = Nothing
55: Set mvarMLSetItems = New MLSetItems
    '扫描语言文件   Scan language files
    bolIsMultiLine = False
    Set oMLSetItem = Nothing
    lngLineNo = 0
    Do While True
ReDo:
        If oTextStream.AtEndOfStream = True Then
            Exit Do
        End If
60:     strItem = oTextStream.ReadLine: lngLineNo = lngLineNo + 1
        '判断是否注释   Judge whether to comment or not
        If Left(strItem, 1) = ";" Then GoTo ReDo:
        '判断是否新段（新文本集）   Determine whether a new paragraph (new text set)
        Dim mvarOptString As String
        If Left(strItem, 1) = "{" Then
            strTmp = gmGetStr(strItem, "{", "}", False)
            If strTmp <> "" Then '新段  New paragraph
                On Error Resume Next
                Set oMLSetItem = mvarMLSetItems.Add(CLng(strTmp), Me)
                If Err.Number <> 0 Then
                    strTmp = "Line:" & lngLineNo & " " & strTmp & "(" & Err.Description & ")" & vbCrLf
                    If mvarOptString = "OK" Then
                        mvarOptString = strTmp
                    Else
                        mvarOptString = mvarOptString & strTmp
                    End If
                    Err.Clear
                End If
                On Error GoTo ErrOcc:
                GoTo ReDo:
            End If
        End If
        If bolIsMultiLine = True Then
            If InStr(strItem, "-->") <> 0 Then
                strTmp = gmGetStr(strItem, "", "-->")
                strValue = strValue & strTmp
                bolIsMultiLine = False
                If Not (oMLSetItem Is Nothing) Then
                    On Error Resume Next
                    oMLSetItem.MLItems.Add strKey, strValue
                    If Err.Number <> 0 Then
                        strTmp = "Line:" & lngLineNo & " " & strKey & "(" & Err.Description & ")" & vbCrLf
                        If mvarOptString = "OK" Then
                            mvarOptString = strTmp
                        Else
                            mvarOptString = mvarOptString & strTmp
                        End If
                        Err.Clear
                    End If
                    On Error GoTo ErrOcc:
                End If
                GoTo ReDo:
            Else
                strValue = strValue & strItem & vbCrLf
            End If
        Else '不是读多行
'            strItem = LTrim(strItem)
            strTmp = gmGetStr(strItem, "[", "]=", True)
            If strTmp <> "" Then '发现元素
                '判断是否要求多行
                strKey = strTmp
                If Left(strItem, 4) = "<!--" Then
                    strValue = gmGetStr(strItem, "<!--", "-->")
                    If strValue = "" Then
                        strValue = gmGetStr(strItem, "<!--", "") & vbCrLf
                        bolIsMultiLine = True
                        GoTo ReDo:
                    End If
                End If
                If Not (oMLSetItem Is Nothing) Then
                    strValue = strItem
                    On Error Resume Next
                    oMLSetItem.MLItems.Add strKey, strValue
                    If Err.Number <> 0 Then
                        strTmp = "Line:" & lngLineNo & " " & strKey & "(" & Err.Description & ")" & vbCrLf
                        If mvarOptString = "OK" Then
                            mvarOptString = strTmp
                        Else
                            mvarOptString = mvarOptString & strTmp
                        End If
                        Err.Clear
                    End If
                    On Error GoTo ErrOcc:
                End If
                GoTo ReDo:
            End If
        End If
    Loop
    'OK
    If mvarOptString <> "OK" Then
        Err.Raise -1, , mvarOptString
    End If
    
    moPigLog.ClearErr
    Set oFile = Nothing
    Set oFolder = Nothing
    Set oTextStream = Nothing
    Set oMLSetItem = Nothing
    On Error GoTo 0
    Exit Sub
ErrOcc:
    moPigLog.SetSubErrInf "ChangLang", "", Err
    Set oFile = Nothing
    Set oFolder = Nothing
    Set oTextStream = Nothing
    Set oMLSetItem = Nothing
    On Error GoTo 0
End Sub



Public Function ShowAllLangInf() As String
    Dim oMLSetItem As MLSetItem
    ShowAllLangInf = "语言标识" & vbTab & "语言信息" & vbCrLf
    For Each oMLSetItem In Me.MLSetItems
        With oMLSetItem
            ShowAllLangInf = ShowAllLangInf & CStr(.LangID) & vbTab & .LocalLangInf & vbCrLf
        End With
    Next
End Function


Public Function GetText(ByVal MLSetItemKey As String, ByVal MLItemKey As String, Optional ByVal DefaultValue As String = "") As String

    On Error Resume Next
    GetText = mvarMLSetItems(MLSetItemKey).MLItems(MLItemKey).Value
    If Err.Number <> 0 Then
        GetText = DefaultValue
        Err.Clear
    End If
    On Error GoTo 0
End Function



