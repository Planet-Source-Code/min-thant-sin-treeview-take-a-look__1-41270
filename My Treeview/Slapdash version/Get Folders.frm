VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop Populating"
      Height          =   390
      Left            =   5625
      TabIndex        =   5
      Top             =   5925
      Width           =   2265
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "&Populate Treeview Now!"
      Default         =   -1  'True
      Height          =   390
      Left            =   3225
      TabIndex        =   4
      Top             =   5925
      Width           =   2265
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4590
      Left            =   3225
      TabIndex        =   2
      Top             =   75
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   8096
      _Version        =   393217
      Indentation     =   273
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.DirListBox Dir1 
      Height          =   5940
      Left            =   75
      TabIndex        =   1
      Top             =   375
      Width           =   3090
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3090
   End
   Begin VB.Label lblPath 
      BorderStyle     =   1  'Fixed Single
      Height          =   1065
      Left            =   3225
      TabIndex        =   3
      Top             =   4725
      Width           =   4665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Slipshod, slapdash version

Private RootDir As String
Private FolderName As String
Private boolCancel As Boolean

Private Sub cmdPopulate_Click()
      boolCancel = False
      TreeView1.Nodes.Clear
      RootDir = Dir1.Path
      
      Dim RootNode As Node
      
      FolderName = GetDirName(Dir1.Path)
      Set RootNode = TreeView1.Nodes.Add(, , RootDir, FolderName)
      Call TraverseDirs(RootDir, RootNode)
End Sub

Sub TraverseDirs(Path As String, ParentNode As Node)
      Dim X As Long
      Dim NewKey As String
      Dim KidNode As Node
      
      If Not boolCancel Then
            For X = 0 To Dir1.ListCount - 1
                  Dir1.Path = Dir1.List(X)
                  FolderName = GetDirName(Dir1.Path)
                  NewKey = AddSlash(Path) & FolderName
                  Set KidNode = TreeView1.Nodes.Add(ParentNode, tvwChild, NewKey, FolderName)
                  Call TraverseDirs(Dir1.Path, KidNode)
                  DoEvents
            Next X
      End If
      
      If Dir1.List(-1) <> RootDir Then Dir1.Path = Dir1.List(-2)
End Sub

Function AddSlash(Path As String) As String
      AddSlash = IIf(Right(Path, 1) = "\", Path, Path & "\")
End Function

Function GetDirName(Path As String) As String
      Dim X As Integer
      
      If Len(Path) <= 3 Then
            If Right(Path, 1) = "\" Then
                  GetDirName = Left$(Path, 2)
            End If
      Else
            For X = Len(Path) To 1 Step -1
                  If Mid(Path, X, 1) = "\" Then
                        GetDirName = Mid$(Path, X + 1, Len(Path) - X)
                        Exit Function
                  End If
            Next X
      End If
End Function

Private Sub cmdStop_Click()
      boolCancel = True
End Sub

Private Sub Dir1_Change()
      ChDir Dir1.Path
End Sub

Private Sub Drive1_Change()
      ChDrive Drive1.Drive
      Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Unload(Cancel As Integer)
      boolCancel = True
      End
End Sub
