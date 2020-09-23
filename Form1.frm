VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4965
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel "
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwFileList 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7435
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   2040
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFileInfo 
      Caption         =   "&Get Version Info"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelectDir 
      Caption         =   "&Select Directory"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4605
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsSearchParam 
         Caption         =   "&Set Search Parameters"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Sub Command1_Click()
blnExit = True
End Sub

Private Sub cmdFileInfo_Click()
Dim rs As VbMsgBoxResult
mnuFileExit.Enabled = False
    If strDirName = "" Then
        rs = MsgBox("Please select a directory!", vbOKOnly, "No directory selected")
        Exit Sub
    End If
    Get_File_Version (strDirName)

End Sub

Private Sub cmdSelectDir_Click()

    Browse_Folder (strDirName)
    Label2.Caption = strDirName
    File1.Path = strDirName

End Sub

Private Sub Form_Load()

  Dim colX(2) As ColumnHeader
  Dim i As Integer

    Label1.Caption = "Directory to check: "

    lvwFileList.View = lvwReport

    Set colX(1) = lvwFileList.ColumnHeaders.Add(, , "FileName")
    Set colX(2) = lvwFileList.ColumnHeaders.Add(, , "Version")
    colX(1).Width = (lvwFileList.Width / 2) + 10
    colX(2).Width = (lvwFileList.Width / 2) + 10
    'blnExit = False
    strDirName = App.Path
End Sub


Private Sub Form_Unload(Cancel As Integer)
blnExit = True

End Sub

Private Sub lvwFileList_Click()
'if vbbuttons
End Sub

Private Sub lvwFileList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If lvwFileList.SortKey = (ColumnHeader.Index - 1) Then
    If lvwFileList.SortOrder = lvwAscending Then
        lvwFileList.SortOrder = lvwDescending
    Else
        lvwFileList.SortOrder = lvwAscending
    End If
Else
lvwFileList.SortKey = ColumnHeader.Index - 1
lvwFileList.Refresh

End If
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuOptionsSearchParam_Click()
MsgBox "Sorry, not implemented yet"
End Sub
