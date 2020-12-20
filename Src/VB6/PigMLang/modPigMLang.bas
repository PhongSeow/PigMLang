Attribute VB_Name = "modPigMLang"
'************************************************
'Class name: Declaration
'Author: Seow Phong
'Organization: Seow Phong Studio(http://en.seowphong.com)
'Description: objectified File
'Version: 1.0.1
'Created: July 16, 2020
'************************************************
Option Explicit

Public Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Public Const LOCALE_SLANGUAGE = &H2         '  localized name of language
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public goFS As pFileSystemObject

Public Function GetObjPropertyText(oCtl As Control, Optional IsFont As Boolean = True) As String
Dim lngIndex As Long, strText As String, bolTmp As Boolean, strCtlName As String

    On Error Resume Next
    '清除错误对象
    If Err.Number <> 0 Then Err.Clear
    '判断控件是否可见
    bolTmp = oCtl.Visible
    If Err.Number = 0 Then
        GetObjPropertyText = ""
        '有界面
        '判断是否控件数组
        lngIndex = oCtl.Index
        If Err.Number <> 0 Then
            '不是控件数组
            lngIndex = -1
            Err.Clear
        End If
        '取出控件名称
        strCtlName = oCtl.Name
        If lngIndex > -1 Then strCtlName = strCtlName & "(" & CStr(lngIndex) & ")" '控件数组加上维数
        '判断是否有标题
        strText = "": strText = oCtl.Caption
        If Err.Number = 0 Then '有
            GetObjPropertyText = GetObjPropertyText & "[" & strCtlName & "_Caption]="
            If InStr(strText, vbCrLf) <> 0 Then '有回车符
                GetObjPropertyText = GetObjPropertyText & "<!--" & strText & "-->" & vbCrLf
            Else
                GetObjPropertyText = GetObjPropertyText & strText & vbCrLf
            End If
        Else
            Err.Clear
        End If
        '判断是否有提示文本
        strText = "": strText = oCtl.ToolTipText
        If Err.Number = 0 Then '有
            GetObjPropertyText = GetObjPropertyText & "[" & strCtlName & "_ToolTipText]="
            If InStr(strText, vbCrLf) <> 0 Then '有回车符
                GetObjPropertyText = GetObjPropertyText & "<!--" & strText & "-->" & vbCrLf
            Else
                GetObjPropertyText = GetObjPropertyText & strText & vbCrLf
            End If
        Else
            Err.Clear
        End If
        If IsFont = True Then
            '判断是否有字体
            strText = "": strText = oCtl.Font.Name
            If Err.Number = 0 Then '有
                GetObjPropertyText = GetObjPropertyText & "[" & strCtlName & "_FontName]=" & strText & vbCrLf
                GetObjPropertyText = GetObjPropertyText & "[" & strCtlName & "_FontSize]=" & oCtl.Font.Size & vbCrLf
            Else
                Err.Clear
            End If
            strText = "": strText = oCtl.CellFontName
            If Err.Number = 0 Then '有
                GetObjPropertyText = GetObjPropertyText & "[" & strCtlName & "_CellFontName]=" & strText & vbCrLf
                GetObjPropertyText = GetObjPropertyText & "[" & strCtlName & "_CellFontSize]=" & oCtl.CellFontSize & vbCrLf
            Else
                Err.Clear
            End If
        End If
    End If
End Function

Public Sub SetObjPropertyText(oCtl As Control, oMLSetItem As MLSetItem, Optional IsFont As Boolean = True)
'更新控件的属性文本
Dim lngIndex As Long, strText As String, bolTmp As Boolean, strCtlName As String
Dim strKey As String, strValue As String, strTypeName As String, i As Long

    On Error Resume Next
    '清除错误对象
    If Err.Number <> 0 Then Err.Clear
    '判断控件是否可见
    bolTmp = oCtl.Visible
    If Err.Number = 0 Then
        '有界面
        '判断是否控件数组
        lngIndex = oCtl.Index
        If Err.Number <> 0 Then
            '不是控件数组
            lngIndex = -1
            Err.Clear
        End If
        
        '确定控件的名称及类型
        strCtlName = oCtl.Name
        strTypeName = TypeName(oCtl)
        If lngIndex > -1 Then strCtlName = strCtlName & "(" & CStr(lngIndex) & ")" '控件数组加上维数
        '更新控件标题
        strKey = strCtlName & "_Caption"
        strValue = oMLSetItem.MLItems(strKey).Value
        If Err.Number <> 0 Then
            Err.Clear
        Else
            oCtl.Caption = strValue
            If Err.Number <> 0 Then Err.Clear
        End If
        '更新提示文本
        strKey = strCtlName & "_ToolTipText"
        strValue = oMLSetItem.MLItems(strKey).Value
        If Err.Number <> 0 Then
            Err.Clear
        Else
            oCtl.ToolTipText = strValue
            If Err.Number <> 0 Then Err.Clear
        End If
        If IsFont = True Then
            '更新字体
            strKey = strCtlName & "_FontName"
            strValue = oMLSetItem.MLItems(strKey).Value
            If Err.Number <> 0 Then
                Err.Clear
            Else
                oCtl.Font.Name = strValue
                If Err.Number <> 0 Then Err.Clear
            End If
            strKey = strCtlName & "_FontSize"
            strValue = oMLSetItem.MLItems(strKey).Value
            If Err.Number <> 0 Then
                Err.Clear
            Else
                oCtl.Font.Size = CCur(strValue)
                If Err.Number <> 0 Then Err.Clear
            End If
            strKey = strCtlName & "_CellFontName"
            strValue = oMLSetItem.MLItems(strKey).Value
            If Err.Number <> 0 Then
                Err.Clear
            Else
                oCtl.Font.Name = strValue
                If Err.Number <> 0 Then Err.Clear
            End If
            strKey = strCtlName & "_CellFontSize"
            strValue = oMLSetItem.MLItems(strKey).Value
            If Err.Number <> 0 Then
                Err.Clear
            Else
                oCtl.Font.Size = CCur(strValue)
                If Err.Number <> 0 Then Err.Clear
            End If
        End If
        '处理字控件集
        Select Case strTypeName
        Case "StatusBar"
            For i = 1 To oCtl.Panels.Count
                With oCtl.Panels(i)
                    strKey = strCtlName & "_Panels(" & i & ")_ToolTipText"
                    strValue = oMLSetItem.MLItems(strKey).Value
                    If Err.Number <> 0 Then
                        Err.Clear
                    Else
                        .ToolTipText = strValue
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End With
            Next
        Case "CoolBar"
            For i = 1 To oCtl.Bands.Count
                With oCtl.Bands(i)
                    strKey = strCtlName & "_Panels(" & i & ")_Caption"
                    strValue = oMLSetItem.MLItems(strKey).Value
                    If Err.Number <> 0 Then
                        Err.Clear
                    Else
                        .Caption = strValue
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End With
            Next
        Case "TabStrip"
            For i = 1 To oCtl.Tabs.Count
                With oCtl.Tabs(i)
                    strKey = strCtlName & "_Tabs(" & i & ")_Caption"
                    strValue = oMLSetItem.MLItems(strKey).Value
                    If Err.Number <> 0 Then
                        Err.Clear
                    Else
                        .Caption = strValue
                        If Err.Number <> 0 Then Err.Clear
                    End If
                    strKey = strCtlName & "_Tabs(" & i & ")_ToolTipText"
                    strValue = oMLSetItem.MLItems(strKey).Value
                    If Err.Number <> 0 Then
                        Err.Clear
                    Else
                        .ToolTipText = strValue
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End With
            Next
        Case "Toolbar"
            For i = 1 To oCtl.Buttons.Count
                With oCtl.Buttons(i)
                    strKey = strCtlName & "_Buttons(" & i & ")_Caption"
                    strValue = oMLSetItem.MLItems(strKey).Value
                    If Err.Number <> 0 Then
                        Err.Clear
                    Else
                        .Caption = strValue
                        If Err.Number <> 0 Then Err.Clear
                    End If
                    strKey = strCtlName & "_Buttons(" & i & ")_ToolTipText"
                    strValue = oMLSetItem.MLItems(strKey).Value
                    If Err.Number <> 0 Then
                        Err.Clear
                    Else
                        .ToolTipText = strValue
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End With
            Next
        End Select
    End If
End Sub

Public Function gmGetStr(SourceStr As String, strBegin As String, strEnd As String, Optional IsCut As Boolean = True) As String
Dim lngBegin As Long
Dim lngEnd As Long
Dim lngBeginLen As Long
Dim lngEndLen As Long
    
    On Error GoTo ErrOcc:
    lngBeginLen = Len(strBegin)
    lngBegin = InStr(SourceStr, strBegin)
    lngEndLen = Len(strEnd)
    If lngEndLen = 0 Then
        lngEnd = Len(SourceStr) + 1
    Else
        lngEnd = InStr(lngBegin + lngBeginLen + 1, SourceStr, strEnd): If lngBegin = 0 Then GoTo ErrOcc:
    End If
    If lngEnd <= lngBegin Then GoTo ErrOcc:
    If lngBegin = 0 Then GoTo ErrOcc:
    
    gmGetStr = Mid(SourceStr, lngBegin + lngBeginLen, (lngEnd - lngBegin - lngBeginLen))
    If IsCut = True Then
        SourceStr = Left(SourceStr, lngBegin - 1) & Mid(SourceStr, lngEnd + lngEndLen)
    End If
    On Error GoTo 0
    Exit Function
    
ErrOcc:
    gmGetStr = ""
    On Error GoTo 0
End Function

Public Function gmGetLangInf(ByVal LangID As Long) As String
Dim strData As String * 256, lngRet As Long

    lngRet = GetLocaleInfo(LangID, LOCALE_SLANGUAGE, strData, 256)
    If lngRet = 0 Then
        gmGetLangInf = ""
    Else
        strData = Replace(strData, vbNullChar, vbNullString)
        gmGetLangInf = Trim(strData)
    End If
    
End Function

Public Function gmGetAboutLangIDList(LangID As Long) As String
Const MAX_I As Integer = 32
Dim alanglist(1 To MAX_I) As String, i As Integer, strTmp As String

    alanglist(1) = "<1><1025><2049><3073><4097><5121><6145><7169><8193><9217><10241><11265><12289><13313><14337><15361><16385>"
    alanglist(2) = "<1024><2048><2052><4100>"   '简体中文
    alanglist(3) = "<9><1033><2057><3081><4105><5129><6153><7177><8201><9225><10249><11273><12297><13321>" '英语
    alanglist(4) = "<12><1036><2060><3084><4108><5132><6156>"
    alanglist(5) = "<2><1026>"
    alanglist(6) = "<5><1029>"
    alanglist(7) = "<11><1035>"
    alanglist(8) = "<14><1038>"
    alanglist(9) = "<16><1040><2064>"
    alanglist(10) = "<18><1042>"
    alanglist(11) = "<20><2068><1044>"
    alanglist(12) = "<22><1046><2070>"
    alanglist(13) = "<25><1049>"
    alanglist(14) = "<27><1051>"
    alanglist(15) = "<29><1053><2077>"
    alanglist(16) = "<44><1068><2092>"
    alanglist(17) = "<47><1071>"
    alanglist(18) = "<55><1079>"
    alanglist(19) = "<57><1081>"
    alanglist(20) = "<63><1087>"
    alanglist(21) = "<67><1091>"
    alanglist(22) = "<73><1097>"
    alanglist(23) = "<79><1103>"
    alanglist(24) = "<2074><3098>"
    alanglist(25) = "<30><1054>"
    alanglist(26) = "<32><1056>"
    alanglist(27) = "<34><1058>"
    alanglist(28) = "<36><1060>"
    alanglist(29) = "<39><1063>"
    alanglist(30) = "<38><1062>"
    alanglist(31) = "<41><1065>"
    alanglist(32) = "<3076><1028><4><5124>"   '繁体中文
    
    strTmp = "<" & LangID & ">"
    gmGetAboutLangIDList = ""
    For i = 1 To MAX_I
        If InStr(alanglist(i), strTmp) <> 0 Then
            gmGetAboutLangIDList = alanglist(i)
            Exit For
        End If
    Next
    
End Function

'https://blog.csdn.net/lqdjdy/article/details/1915442
'0x0000 Language Neutral
'0x0400 Process Default Language
'0x0401 Arabic(Saudi Arabia)
'0x0801 Arabic(Iraq)
'0x0c01 Arabic(Egypt)
'0x1001 Arabic(Libya)
'0x1401 Arabic(Algeria)
'0x1801 Arabic(Morocco)
'0x1c01 Arabic(Tunisia)
'0x2001 Arabic(Oman)
'0x2401 Arabic(Yemen)
'0x2801 Arabic(Syria)
'0x2c01 Arabic(Jordan)
'0x3001 Arabic(Lebanon)
'0x3401 Arabic(Kuwait)
'0x3801 Arabic(U.A.E.)
'0x3c01 Arabic(Bahrain)
'0x4001 Arabic(Qatar)
'0x0402 Bulgarian
'0x0403 Catalan
'0x0404 Chinese(Taiwan Region)
'0x0804 Chinese(PRC)
'0x0c04 Chinese(Hong Kong SAR, PRC)
'0x1004 Chinese(Singapore)
'0x0405 Czech
'0x0406 Danish
'0x0407 German(Standard)
'0x0807 German(Swiss)
'0x0c07 German(Austrian)
'0x1007 German(Luxembourg)
'0x1407 German(Liechtenstein)
'0x0408 Greek
'0x0409 English(United States)
'0x0809 English(United Kingdom)
'0x0c09 English(Australian)
'0x1009 English(Canadian)
'0x1409 English(New Zealand)
'0x1809 English(Ireland)
'0x1c09 English(South Africa)
'0x2009 English(Jamaica)
'0x2409 English(Caribbean)
'0x2809 English(Belize)
'0x2c09 English(Trinidad)
'0x040a Spanish(Traditional Sort)
'0x080a Spanish(Mexican)
'0x0c0a Spanish(Modern Sort)
'0x100a Spanish(Guatemala)
'0x140a Spanish(Costa Rica)
'0x180a Spanish(Panama)
'0x1c0a Spanish(Dominican Republic)
'0x200a Spanish(Venezuela)
'0x240a Spanish(Colombia)
'0x280a Spanish(Peru)
'0x2c0a Spanish(Argentina)
'0x300a Spanish(Ecuador)
'0x340a Spanish(Chile)
'0x380a Spanish(Uruguay)
'0x3c0a Spanish(Paraguay)
'0x400a Spanish(Bolivia)
'0x440a Spanish(El Salvador)
'0x480a Spanish(Honduras)
'0x4c0a Spanish(Nicaragua)
'0x500a Spanish(Puerto Rico)
'0x040b Finnish
'0x040c French(Standard)
'0x080c French(Belgian)
'0x0c0c French(Canadian)
'0x100c French(Swiss)
'0x140c French(Luxembourg)
'0x040d Hebrew
'0x040e Hungarian
'0x040f Icelandic
'0x0410 Italian(Standard)
'0x0810 Italian(Swiss)
'0x0411 Japanese
'0x0412 Korean
'0x0812 Korean(Johab)
'0x0413 Dutch(Standard)
'0x0813 Dutch(Belgian)
'0x0414 Norwegian(Bokmal)
'0x0814 Norwegian(Nynorsk)
'0x0415 Polish
'0x0416 Portuguese(Brazilian)
'0x0816 Portuguese(Standard)
'0x0418 Romanian
'0x0419 Russian
'0x041a Croatian
'0x081a Serbian(Latin)
'0x0c1a Serbian(Cyrillic)
'0x041b Slovak
'0x041c Albanian
'0x041d Swedish
'0x081d Swedish(Finland)
'0x041e Thai
'0x041f Turkish
'0x0421 Indonesian
'0x0422 Ukrainian
'0x0423 Belarusian
'0x0424 Slovenian
'0x0425 Estonian
'0x0426 Latvian
'0x0427 Lithuanian
'0x0429 Farsi
'0x042a Vietnamese
'0x042d Basque
'0x0436 Afrikaans
'0x0438 Faeroese
