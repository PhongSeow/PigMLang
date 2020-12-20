VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'************************************************
'Class name: Obj Folder
'Author: Seow Phong
'Organization: Seow Phong Studio(http://en.seowphong.com)
'Description: objectified Folder
'Version: 1.0.3
'Created: July 16, 2020
'1.0.2 2020-7-19 Class_ Terminate
'1.0.3 September 1, 2020 Add IsObjReady
'************************************************
Option Explicit

Public Obj As Object
'Public Obj As Scripting.Folder

Public Property Get Name() As String
    Name = Me.Obj.Name
End Property

Public Property Get DateLastModified() As Date
    DateLastModified = Me.Obj.DateLastModified
End Property

Public Property Get DateCreated() As Date
    DateCreated = Me.Obj.DateCreated
End Property


Public Property Get Path() As String
    Path = Me.Obj.Path
End Property

Public Property Get ShortName() As String
    ShortName = Me.Obj.ShortName
End Property

Public Property Get ShortPath() As String
    ShortPath = Me.Obj.ShortPath
End Property

Public Property Get Size() As Long
    Size = Me.Obj.Size
End Property

Public Property Get Files() As Object
    Set Files = Me.Obj.Files
End Property

Public Property Get ParentFolder() As Object
    Set ParentFolder = Me.Obj.ParentFolder
End Property

Public Property Get IsRootFolder() As Boolean
    IsRootFolder = Me.Obj.IsRootFolder
End Property

Public Property Get SubFolders() As Object
    Set SubFolders = Me.Obj.SubFolders
End Property

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

