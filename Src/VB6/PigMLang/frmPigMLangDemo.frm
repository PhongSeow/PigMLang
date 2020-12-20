VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPigMLangDemo 
   Caption         =   "PigMLang"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   10080
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtMain 
      Height          =   2535
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMultiLangApp 
      Caption         =   "MultiLangApp"
      Begin VB.Menu mnuMultiLangApp_MKDictString 
         Caption         =   "MKDictString"
         Begin VB.Menu mnuMultiLangApp_MKDictString_Me 
            Caption         =   "Me"
         End
         Begin VB.Menu mnuMultiLangApp_MKDictString_frmTestMLang 
            Caption         =   "frmTestMLang"
         End
      End
      Begin VB.Menu mnuMultiLangApp_ShowProperty 
         Caption         =   "ShowProperty"
      End
      Begin VB.Menu mnuMultiLangApp_LoadDictFile 
         Caption         =   "LoadDictFile"
      End
      Begin VB.Menu mnuMultiLangApp_ShowAllLangInf 
         Caption         =   "ShowAllLangInf"
      End
      Begin VB.Menu mnupFileSystemObject_CreateFolder 
         Caption         =   "CreateFolder"
      End
   End
   Begin VB.Menu mnupTextStream 
      Caption         =   "pTextStream"
      Begin VB.Menu mnupTextStream_ReadFile 
         Caption         =   "Read File"
      End
      Begin VB.Menu mnupTextStream_WriteFile 
         Caption         =   "Write File"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp_OnlineDoc 
         Caption         =   "Online documentation"
      End
      Begin VB.Menu mnuHelp_SPS 
         Caption         =   "Seow Phong Studio"
      End
   End
End
Attribute VB_Name = "frmPigMLangDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oMultiLangApp As MultiLangApp

Private Sub Form_Load()
    Me.Caption = App.ProductName & App.FileDescription
    Set oMultiLangApp = New MultiLangApp
'    oMultiLangApp.LoadDictFile App.Path & "\Lang", App.Title
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With Me.txtMain
        .Top = 50
        .Left = 50
        .Width = Me.ScaleWidth - 100
        .Height = Me.ScaleHeight - 100
    End With
    On Error GoTo 0
End Sub

Private Sub mnuFile_Exit_Click()
    Unload Me
End Sub


Private Sub mnuHelp_OnlineDoc_Click()
    Shell "explorer https://en.seowphong.com/oss/PigObjFs/"
End Sub

Private Sub mnuHelp_SPS_Click()
    Shell "explorer https://en.seowphong.com"
End Sub

Private Sub mnuMultiLangApp_LoadDictFile_Click()
    Dim strRet As String, strFilePath As String, strFileTitle As String, strText As String
    With Me.cdMain
        .Flags = 0
        .DialogTitle = mnuMultiLangApp_LoadDictFile.Caption
        .InitDir = App.Path
        .Filter = "多语言文件" & "(*.ml)|" & "所有文件" & "(*.*)|*.*"
        .FileName = ""
        .ShowOpen
        If .FileName = "" Or .Flags = 0 Then
            Exit Sub
        Else
            strFilePath = .FileName
        End If
    End With
    
    With oMultiLangApp
         .LoadDictFile strFilePath, strFileTitle
        If .LastErr <> "" Then MsgBox .LastErr, vbCritical, mnuMultiLangApp_LoadDictFile.Caption
        strText = strText & "CurrentLangID=" & .CurrentLangID & vbCrLf
        strText = strText & "BindFormList=" & .BindFormList & vbCrLf
        strText = strText & "LangIDList=" & .LangIDList & vbCrLf
        strText = strText & "LangList=" & .LangList & vbCrLf
    End With
    Me.txtMain.Text = strText
End Sub


Private Sub mnuMultiLangApp_MKDictString_frmTestMLang_Click()
    Me.txtMain.Text = oMultiLangApp.MKDictString(frmTestMLang)
End Sub

Private Sub mnuMultiLangApp_MKDictString_Me_Click()
    Me.txtMain.Text = oMultiLangApp.MKDictString(Me)
End Sub

Private Sub mnuMultiLangApp_ShowAllLangInf_Click()
    Me.txtMain.Text = oMultiLangApp.ShowAllLangInf
End Sub

Private Sub mnuMultiLangApp_ShowProperty_Click()
    Dim strText As String
    With oMultiLangApp
        strText = strText & "CurrentLangID=" & .CurrentLangID & vbCrLf
        strText = strText & "BindFormList=" & .BindFormList & vbCrLf
        strText = strText & "LangIDList=" & .LangIDList & vbCrLf
        strText = strText & "LangList=" & .LangList & vbCrLf
    End With
    Me.txtMain.Text = strText
End Sub
