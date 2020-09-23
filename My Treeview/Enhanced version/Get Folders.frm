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
      Enabled         =   0   'False
      Height          =   390
      Left            =   5625
      TabIndex        =   5
      Top             =   5925
      Width           =   2265
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3450
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Get Folders.frx":0000
            Key             =   "cf1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Get Folders.frx":015A
            Key             =   "of1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Get Folders.frx":02B4
            Key             =   "cf2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Get Folders.frx":084E
            Key             =   "of2"
         EndProperty
      EndProperty
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
      Style           =   7
      ImageList       =   "ImageList1"
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

'Created by Min Thant Sin in 2002
'Thanks to all my teachers. I always use some of their idea.

'///////////////////////////////////////////
'Note: No error handling of DOS dir names such as
'< MyßæFolder > in TraverseDirs sub.
'If the error ( folder doesn't exist ) occurs,
'change the path to the current dir's parent dir
'in the error handling code.
'///////////////////////////////////////////

'Any bugs?
'Feel free to e-mail me at  < minsin999@hotmail.com >

Private OpenIconKey As String      'Key of open folder icon in ImageList1
Private CloseIconKey As String      'Key of close folder icon in ImageList1

Private RootDir As String         'Selected (current) directory in Dir1 control
Private FolderName As String  'Just folder name, not the whole path, used in tree nodes' text
Private boolCancel As Boolean   'Flag to indicate user canceled treeview-populating process

Private Sub cmdPopulate_Click()
      'If the focus is in the Dir1 control and the user
      'presses Enter key, we assume that he/she wants
      'to change the path in that control.
      'Why do I write this code in cmdPopulate click event?
      'Because I set the button's Default property value
      'to True. You can figure it out.
      If Dir1.Path <> Dir1.List(Dir1.ListIndex) Then
            Dir1.Path = Dir1.List(Dir1.ListIndex)
            'Exit sub so the user can take a look at
            'which dir he/she is at
            Exit Sub
      End If
      
      'Init flag
      boolCancel = False
      
      'Disable to avoid repetition of populating treeview
      cmdPopulate.Enabled = False
      
      'Enable because there is an operation taking place
      'and the user can cancel it
      cmdStop.Enabled = True
      cmdStop.SetFocus
      
      'Clear the treeview first
      TreeView1.Nodes.Clear
      
      'Store the currently selected directory
      RootDir = Dir1.Path
      
      Dim RootNode As Node
      
      'Extract a single dir name for tree nodes' text
      FolderName = GetDirName(Dir1.Path)
      
      'Populate treeview with the root node
      Set RootNode = TreeView1.Nodes.Add(, , RootDir, FolderName, CloseIconKey, OpenIconKey)
      
      'Now, get directories under root dir
      Call TraverseDirs(RootDir, RootNode)
      
      '////////////////////////////////////
      'Finished or user canceled operation
      '////////////////////////////////////
      
      'Inform how many dirs found
      MsgBox TreeView1.Nodes.Count & " folders found", vbInformation
      
      'Enable so that the user can populate treeview again
      cmdPopulate.Enabled = True
      
      'Disable because there is no operation to stop
      cmdStop.Enabled = False
      
      'Give the treeview focus
      TreeView1.SetFocus
      
      lblPath.Caption = ""
      
End Sub

Sub TraverseDirs(Path As String, ParentNode As Node)
      Dim X As Long
      Dim NewKey As String    'New key for new node (KidNode)
      Dim KidNode As Node     'Child of the current node (ParentNode)
      
      'If user didn't cancel the operation...
      If Not boolCancel Then
            'Traverse all subdirs under the root dir
            For X = 0 To Dir1.ListCount - 1
            
                  'Change path to each subdir
                  Dir1.Path = Dir1.List(X)
                  
                  'Actually change the path
                  ChDir Dir1.Path
                  
                  'Display which path are we traversing
                  lblPath.Caption = Dir1.Path
                  
                  'Get folder name only, not the whole path
                  FolderName = GetDirName(Dir1.Path)
                  
                  'Create new key
                  NewKey = AddSlash(Path) & FolderName
                 
                 'Populate treeview with the new kid node
                  Set KidNode = TreeView1.Nodes.Add(ParentNode, tvwChild, NewKey, FolderName, CloseIconKey, OpenIconKey)
                  
                  'Traverse recursively
                  Call TraverseDirs(Dir1.Path, KidNode)
                  
                  'Don't hog the processor
                  DoEvents
            Next X
      End If
      
      'Move to parent dir after traversing of each sub dir
      '(-1) is current dir
      '(-2) is current dir's parent dir
      If Dir1.List(-1) <> RootDir Then
            Dir1.Path = Dir1.List(-2)
      End If
End Sub

'This function make sure there is a "\"
'at the end of the Path
Function AddSlash(Path As String) As String
      If Right(Path, 1) = "\" Then
            AddSlash = Path
      Else
            AddSlash = Path & "\"
      End If
End Function

Function GetDirName(Path As String) As String
      Dim X As Integer
      
      'I check for length 3 because Dir1 control doesn't
      'display the drive's name (C:\ [Min Thant Sin]), but
      'just drive letter (C:\)
      If Len(Path) <= 3 Then
            'C:\ or D:\ or E:\ ...etc
            If Right(Path, 1) = "\" Then
                  GetDirName = Left$(Path, 2)   'Returns C: or D: or E: ...etc
            End If
      
      Else 'More than length 3
      
            'Check backward for "\"
            For X = Len(Path) To 1 Step -1
                  'We found it
                  If Mid(Path, X, 1) = "\" Then
                        'Return the dir name only
                        GetDirName = Mid$(Path, X + 1, Len(Path) - X)
                        Exit Function
                  End If
            Next X
            
      End If
End Function

Private Sub cmdStop_Click()
      'I think these are self-explanatory
      boolCancel = True
      cmdPopulate.Enabled = True
      cmdStop.Enabled = False
End Sub

Private Sub Dir1_Change()
      'Change the path
      ChDir Dir1.Path
End Sub

Private Sub Drive1_Change()
      On Error Resume Next
      ChDrive Drive1.Drive
      Dir1.Path = Drive1.Drive
      If Err Then
            Drive1.Drive = Dir1.Path
      End If
End Sub

Private Sub Form_Load()
      'Assign icon keys
      OpenIconKey = "of1"
      CloseIconKey = "cf1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
      boolCancel = True
      End
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
      'Select the collapsed node
      Node.Selected = True
      lblPath.Caption = Node.Key
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
      'This is for the root node
      If Node.Selected Then
            lblPath.Caption = Node.Key
      End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
      'Display currently selected node's key
      lblPath.Caption = Node.Key
End Sub
