VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pFileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'************************************************
'Class name: Obj FileSystemObject
'Author: Seow Phong
'Web site: https://en.seowphong.com
'Description: Objectified file systemobject
'Version: 1.0.3
'Created: July 19, 2020
'1.0.2  August 25, 2020 Add FileExists, FolderExists, CreateFolder
'1.0.3  August 26, 2020 Optimize the CreateFolder to create subdirectories on nonexistent directories
'1.0.4 September 1, 2020 Add IsObjReady
'************************************************
Option Explicit

Private moFS As Object
'Private moFS As Scripting.FileSystemObject

Private Enum xpFilePart
    xpFilePart_Path = 1
    xpFilePart_FileTitle = 2
    xpFilePart_ExtName = 3
    xpFilePart_DriveNo = 4
End Enum

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

Private Sub Class_Initialize()
    Set moFS = CreateObject("Scripting.FileSystemObject")
    Me.InitLog "", ""
End Sub


Private Sub Class_Terminate()
    Set moFS = Nothing
End Sub


Public Sub MoveFile(ByVal Source As String, ByVal Destination As String)
    On Error GoTo ErrOcc:
    moFS.MoveFile Source, Destination
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Sub
ErrOcc:
    moPigLog.SetSubErrInf "MoveFile", "", Err
    On Error GoTo 0
End Sub


Public Sub DeleteFile(ByVal FileSpec As String, Optional ByVal Force As Boolean = False)
    On Error GoTo ErrOcc:
    moFS.DeleteFile FileSpec, Force
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Sub
ErrOcc:
    moPigLog.SetSubErrInf "DeleteFile", "", Err
    On Error GoTo 0
End Sub


Public Function OpenTextFile(ByVal FileName As String, ByVal IOMode As pIOMode, Optional ByVal Create As Boolean = False) As pTextStream
    Dim strStepName As String, strRet As String
    On Error GoTo ErrOcc:
    strStepName = "New pTextStream"
    Set OpenTextFile = New pTextStream
    strStepName = "FS.OpenTextFile"
    Set OpenTextFile.Obj = moFS.OpenTextFile(FileName, IOMode, Create)
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Function
ErrOcc:
    moPigLog.SetSubErrInf "OpenTextFile", strStepName, Err
    On Error GoTo 0
End Function

Public Function GetFile(ByVal FilePath As String) As pFile
    Dim strStepName As String, strRet As String
    On Error GoTo ErrOcc:
    strStepName = "New pFile"
    Set GetFile = New pFile
    strStepName = "FS.GetFile"
    Set GetFile.Obj = moFS.GetFile(FilePath)
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Function
ErrOcc:
    moPigLog.SetSubErrInf "GetFile", strStepName, Err
    On Error GoTo 0
End Function

Public Function GetFolder(ByVal FolderPath As String) As pFolder
    Dim strStepName As String, strRet As String
    On Error GoTo ErrOcc:
    strStepName = "New pFolder"
    Set GetFolder = New pFolder
    strStepName = "FS.GetFolder"
    Set GetFolder.Obj = moFS.GetFolder(FolderPath)
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Function
ErrOcc:
    moPigLog.SetSubErrInf "GetFolder", strStepName, Err
    On Error GoTo 0
End Function



Public Function FolderExists(ByVal FolderSpec As String) As Boolean
    On Error GoTo ErrOcc:
    FolderExists = moFS.FolderExists(FolderSpec)
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Function
ErrOcc:
    moPigLog.SetSubErrInf "FolderExists", "", Err
    FolderExists = False
    On Error GoTo 0
End Function

Public Function FileExists(ByVal FileSpec As String) As Boolean
    On Error GoTo ErrOcc:
    FileExists = moFS.FileExists(FileSpec)
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Function
ErrOcc:
    moPigLog.SetSubErrInf "FileExists", "", Err
    FileExists = False
    On Error GoTo 0
End Function

Public Function CreateFolder(Path As String) As pFolder
    Dim strStepName As String, strRet As String
    On Error GoTo ErrOcc:
    strStepName = "mGEMkDir(" & Path & ")"
    strRet = mGEMkDir(Path)
    If strRet <> "OK" Then Err.Raise -1, , strRet
    strStepName = "New pFolder"
    Set CreateFolder = New pFolder
    strStepName = "FS.GetFolder"
    Set CreateFolder.Obj = moFS.GetFolder(Path)
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Function
ErrOcc:
    moPigLog.SetSubErrInf "CreateFolder", strStepName, Err
    On Error GoTo 0
End Function

Public Sub CopyFile(ByVal Source As String, ByVal Destination As String, Optional ByVal OverWriteFiles As Boolean = True)
    On Error GoTo ErrOcc:
    moFS.CopyFile Source, Destination, OverWriteFiles
    moPigLog.ClearErr
    On Error GoTo 0
    Exit Sub
ErrOcc:
    moPigLog.SetSubErrInf "CopyFile", "", Err
    On Error GoTo 0
End Sub


Private Function mGetFilePart(ByVal FileName As String, Optional FilePart As xpFilePart = xpFilePart_FileTitle) As String
Dim strTemp As String, i As Long, lngLen As Long
Dim strPath As String, strFileTitle As String

    Select Case FilePart
    Case xpFilePart_DriveNo
        mGetFilePart = mGetStr(FileName, "", ":", False)
        If mGetFilePart = "" Then
            mGetFilePart = mGetStr(FileName, "", "$", False)
            If mGetFilePart <> "" Then mGetFilePart = mGetFilePart & "$"
        End If
    Case xpFilePart_ExtName
        'mGetFilePart = mGetStr(FileName, ".", "", False)
        lngLen = Len(FileName)
        For i = lngLen To 1 Step -1
            Select Case Mid(FileName, i, 1)
            Case "/", ":", "$"
                Exit For
            Case "."
                mGetFilePart = Mid(FileName, i + 1)
                Exit For
            End Select
            
        Next
    Case xpFilePart_FileTitle, xpFilePart_Path
        Do While True
            DoEvents
            strTemp = mGetStr(FileName, "", "\", True)
            If Len(strTemp) = 0 Then
                If Right(strPath, 1) = "\" Then
                    If Right(strPath, 2) <> ":\" Then
                        strPath = Left(strPath, Len(strPath) - 1)
                    End If
                ElseIf Left(FileName, 1) = "\" Then
                    strPath = "\"
                    FileName = Mid(FileName, 2)
                End If
                If FileName <> "" Then
                    strFileTitle = FileName
                Else
                    strFileTitle = "."
                End If
                Exit Do
            End If
            strPath = strPath & strTemp & "\"
        Loop
        If FilePart = xpFilePart_FileTitle Then
            mGetFilePart = strFileTitle
        Else
            mGetFilePart = strPath
        End If
    Case Else
        mGetFilePart = "Error"
    End Select
End Function


Private Function mGEMkDir(ByVal Path As String) As String
Dim strDir As String, strDrive As String, strTmp As String
    
    On Error GoTo ErrOcc:
    strDrive = mGetFilePart(Path, xpFilePart_DriveNo)
    If Len(strDrive) = 0 Then
        strDrive = mGetFilePart(CurDir(), xpFilePart_DriveNo) & ":\"
        If Left(Path, 1) = "\" Then
            Path = strDrive & Path
        Else
            Path = CurDir() & "\" & Path
        End If
    Else
        If InStr(strDrive, "$") = 0 Then strDrive = strDrive & ":\"
    End If
    
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    strDir = mGetStr(Path, strDrive, "\", False)
    strDir = strDrive & strDir & "\"
    On Error Resume Next
    Do While True
        DoEvents
        moFS.CreateFolder strDir
        Err.Clear
        strTmp = mGetStr(Path, strDir, "\", False)
        If Len(strTmp) = 0 Then
            Exit Do
        End If
        strDir = strDir & strTmp & "\"
    Loop
    
    mGEMkDir = "OK"
    On Error GoTo 0
    Exit Function
ErrOcc:
    If Err.Number <> 0 Then mGEMkDir = Err.Description
    On Error GoTo 0
End Function

Private Function mGetStr(SourceStr As String, strBegin As String, strEnd As String, Optional IsCut As Boolean = True) As String
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
    
    mGetStr = Mid(SourceStr, lngBegin + lngBeginLen, (lngEnd - lngBegin - lngBeginLen))
    If IsCut = True Then
        SourceStr = Left(SourceStr, lngBegin - 1) & Mid(SourceStr, lngEnd + lngEndLen)
    End If
    On Error GoTo 0
    Exit Function
    
ErrOcc:
    mGetStr = ""
    On Error GoTo 0
End Function

Public Property Get IsObjReady() As Boolean
    If moFS Is Nothing Then
        IsObjReady = False
    Else
        IsObjReady = True
    End If
End Property


