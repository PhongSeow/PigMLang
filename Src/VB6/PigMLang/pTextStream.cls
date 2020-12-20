VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pTextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'************************************************
'Class name: Obj TextStream
'Author: Seow Phong
'Organization: Seow Phong Studio(http://en.seowphong.com)
'Description: objectified TextStream
'Version: 1.0.3
'Created: July 19, 2020
'1.0.2 2020-7-30 add error handling
'1.0.3 September 1, 2020 Add IsObjReady
'************************************************

Option Explicit

Public Obj As Object
'Public Obj As Scripting.TextStream
Private mstrLastErr As String

Public Property Get LastErr() As String
    LastErr = mstrLastErr
End Property

Public Sub CloseObj()
    Obj.Close
End Sub

Public Function AtEndOfLine() As Boolean
    On Error GoTo ErrOcc:
    AtEndOfLine = Obj.AtEndOfLine
    mstrLastErr = ""
    On Error GoTo 0
    Exit Function
ErrOcc:
    mstrLastErr = Err.Description
    On Error GoTo 0
End Function

Public Function AtEndOfStream() As Boolean
    On Error GoTo ErrOcc:
    AtEndOfStream = Obj.AtEndOfStream
    mstrLastErr = ""
    On Error GoTo 0
    Exit Function
ErrOcc:
    mstrLastErr = Err.Description
    On Error GoTo 0
End Function

Public Function ReadAll() As String
    On Error GoTo ErrOcc:
    ReadAll = Obj.ReadAll
    mstrLastErr = ""
    On Error GoTo 0
    Exit Function
ErrOcc:
    mstrLastErr = Err.Description
    On Error GoTo 0
End Function

Public Function ReadLine() As String
    On Error GoTo ErrOcc:
    ReadLine = Obj.ReadLine
    mstrLastErr = ""
    On Error GoTo 0
    Exit Function
ErrOcc:
    mstrLastErr = Err.Description
    On Error GoTo 0
End Function

Public Sub WriteStr(ByVal Text As String)
    On Error GoTo ErrOcc:
    Me.Obj.Write Text
    mstrLastErr = ""
    On Error GoTo 0
    Exit Sub
ErrOcc:
    mstrLastErr = Err.Description
    On Error GoTo 0
End Sub

Public Sub WriteLine(ByVal Text As String)
    On Error GoTo ErrOcc:
    Me.Obj.WriteLine Text
    mstrLastErr = ""
    On Error GoTo 0
    Exit Sub
ErrOcc:
    mstrLastErr = Err.Description
    On Error GoTo 0
End Sub

Public Sub WriteBlankLines(ByVal Lines As Long)
    On Error GoTo ErrOcc:
    Me.Obj.WriteBlankLines Lines
    mstrLastErr = ""
    On Error GoTo 0
    Exit Sub
ErrOcc:
    mstrLastErr = Err.Description
    On Error GoTo 0
End Sub

Public Property Get IsObjReady() As Boolean
    If Me.Obj Is Nothing Then
        IsObjReady = False
    Else
        IsObjReady = True
    End If
End Property


Private Sub Class_Terminate()
    Set Me.Obj = Nothing
End Sub

