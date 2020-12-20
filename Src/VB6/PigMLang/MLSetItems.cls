VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MLSetItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"MLSetItem"
Attribute VB_Ext_KEY = "Member0" ,"MLSetItem"
'*******************************************************
'* Name: MLSetItems
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Multilingual set items
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Created: 11/5/2020
'*******************************************************

Option Explicit

Private mCol As Collection

Friend Function Add(LangID As Long, Parent As MultiLangApp) As MLSetItem
    Dim objNewMember As MLSetItem
    Set objNewMember = New MLSetItem


    With objNewMember
        .LangID = LangID
        Set .Parent = Parent
    End With
    mCol.Add objNewMember, CStr(LangID)


    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As MLSetItem
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    Count = mCol.Count
End Property


Friend Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    Set mCol = Nothing
    On Error GoTo 0
End Sub

