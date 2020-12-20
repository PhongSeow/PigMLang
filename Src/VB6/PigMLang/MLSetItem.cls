VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MLSetItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MLItems"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************
'* Name: MLSetItem
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Multilingual set item
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Created: 11/10/2020
'*******************************************************
Option Explicit

Private mvarLangID As Long   '语言标识   Language identification
Public MLItems As MLItems  '多语言信息集    Multilingual information set
Public Parent As MultiLangApp '父对象   Parent object
Private mvarLocalLangInf As String '本地语言显示的语言和国家或地区名称 The name of the language and country or region displayed in the local language


Public Property Get LocalLangInf() As String
    If mvarLocalLangInf = "" Then
        mvarLocalLangInf = gmGetLangInf(mvarLangID)
    End If
    LocalLangInf = mvarLocalLangInf
End Property

Friend Property Let LangID(ByVal vData As Long)
    mvarLangID = vData
End Property


Public Property Get LangID() As Long
    LangID = mvarLangID
End Property



Private Sub Class_Initialize()
    Set Me.MLItems = New MLItems
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set Me.MLItems = Nothing
    On Error GoTo 0
End Sub
