VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PigLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'************************************************
'Class Name: Log processing class
'Author: Seow Phong
'Organization: Seow Phong Studio(http://en.seowphong.com)
'Description: Log processing
'Version: 1.0.3
'Created: May 26, 2020
'1.0.2  August 25, 2020 Added GetFuncDemo
'1.0.3  August 26, 2020 Make LogDir and LogFileTitle properties writable
'************************************************
Option Explicit

Private mvarLastErr As String 'Recent errors
Private mstrClassName As String 'Class name
Private mstrLogFileTitle As String 'Log file title
Private mstrLogDir As String 'Log directory
Private mstrLogFilePath As String 'log file

Private Sub mOptLogInf(ByVal OptString As String, Optional ByVal LogFileName As String, Optional IsShowTime As Boolean = True)
Dim intFreeFile As Integer
    On Error Resume Next
    intFreeFile = FreeFile()
    Open LogFileName For Append Shared As #intFreeFile
    If IsShowTime = True Then
        Print #intFreeFile, Format(Now, "yyyy-mm-dd hh:mm:ss")
    End If
    Print #intFreeFile, OptString
    Close #intFreeFile
    On Error GoTo 0
End Sub


Public Sub Init(ByVal ClassName As String, Optional ByVal LogDir As String = "", Optional ByVal LogFileTitle As String = "")
    mstrClassName = ClassName
    If LogFileTitle = "" Then LogFileTitle = App.Title
    mstrLogFileTitle = LogFileTitle
    If LogDir = "" Then LogDir = App.Path
    mstrLogDir = LogDir
End Sub

Public Sub PrintSubErrInf(SubName As String, StepName As String, ByRef exIn As ErrObject)
    Dim strLogInf As String
    strLogInf = Me.GetSubErrInf(SubName, StepName, exIn)
    mOptLogInf strLogInf, mstrLogFilePath
End Sub


Public Sub RefreshFilePath()
    Dim strYyyyMmDd As String
    strYyyyMmDd = Format(Now, "yyyymmdd")
    mstrLogFilePath = mstrLogDir & "\" & mstrLogFileTitle & strYyyyMmDd & ".log"
    mstrLogFilePath = mstrLogDir & "\" & mstrLogFileTitle & strYyyyMmDd & ".err"
End Sub

Public Sub ClearErr()
    mvarLastErr = ""
End Sub

Public Property Get LastErr() As String
    LastErr = mvarLastErr
End Property

Public Property Get LogDir() As String
    LogDir = mstrLogDir
End Property

Public Property Let LogDir(ByVal vData As String)
    mstrLogDir = vData
    Me.RefreshFilePath
End Property

Public Property Let LogFileTitle(ByVal vData As String)
    mstrLogFileTitle = vData
    Me.RefreshFilePath
End Property


Public Sub SetSubErrInf(SubName As String, StepName As String, ByRef exIn As ErrObject)
    mvarLastErr = Me.GetSubErrInf(SubName, StepName, exIn)
End Sub


Public Function GetSubErrInf(SubName As String, StepName As String, ByRef exIn As ErrObject) As String
    GetSubErrInf = mstrClassName & "." & SubName & "[" & StepName & "]" & exIn.Description
End Function


Public Function GetFuncDemo(ByVal FuncName As String, ByVal FuncPara As String) As String
    GetFuncDemo = "Public Function " & FuncName & "(" & FuncPara & ") As String" & vbCrLf
    GetFuncDemo = GetFuncDemo & vbTab & "Dim strStepName As String" & vbCrLf
    GetFuncDemo = GetFuncDemo & vbTab & "On Error GoTo ErrOcc:" & vbCrLf
    GetFuncDemo = GetFuncDemo & vbTab & FuncName & " = ""OK""" & vbCrLf
    GetFuncDemo = GetFuncDemo & vbTab & "On Error GoTo 0" & vbCrLf
    GetFuncDemo = GetFuncDemo & vbTab & "Exit Function" & vbCrLf
    GetFuncDemo = GetFuncDemo & "ErrOcc:" & vbCrLf
    GetFuncDemo = GetFuncDemo & vbTab & FuncName & " = moPigLog.GetSubErrInf(""" & FuncName & """,strStepName,Err)" & vbCrLf
    GetFuncDemo = GetFuncDemo & vbTab & "On Error GoTo 0" & vbCrLf
    GetFuncDemo = GetFuncDemo & "End Function" & vbCrLf
End Function

Public Function GetSubDemo(ByVal SubName As String, ByVal SubPara As String) As String
    GetSubDemo = "Public Sub " & SubName & "(" & SubPara & ") As String" & vbCrLf
    GetSubDemo = GetSubDemo & vbTab & "Dim strStepName As String" & vbCrLf
    GetSubDemo = GetSubDemo & vbTab & "On Error GoTo ErrOcc:" & vbCrLf
    GetSubDemo = GetSubDemo & vbTab & "moPigLog.ClearErr" & vbCrLf
    GetSubDemo = GetSubDemo & vbTab & "On Error GoTo 0" & vbCrLf
    GetSubDemo = GetSubDemo & vbTab & "Exit Sub" & vbCrLf
    GetSubDemo = GetSubDemo & "ErrOcc:" & vbCrLf
    GetSubDemo = GetSubDemo & vbTab & "moPigLog.SetSubErrInf(""" & SubName & """,strStepName,Err)" & vbCrLf
    GetSubDemo = GetSubDemo & vbTab & "On Error GoTo 0" & vbCrLf
    GetSubDemo = GetSubDemo & "End Sub" & vbCrLf
End Function



