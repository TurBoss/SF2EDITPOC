VERSION 5.00
Begin VB.Form SelectFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select File"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6645
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Set As Default Path"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   5640
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   5355
      Left            =   2760
      Pattern         =   "*.smd;*.bin;sf2edit.conf"
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   5265
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "SelectFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdDefault_Click()
  Dim Freedfile As Long
  
  Freedfile = FreeFile()
  
  Open App.Path & "/Data/Defaultpath.txt" For Output As #Freedfile
   Print #Freedfile, Dir1.Path
  Close #Freedfile
  
End Sub

Private Sub cmdOpen_Click()

Dim FileAbsolutePath As String
    
    If File1.FileName = SF2EDITCONF_FILENAME Then
       disasmMode = True
    Else
       disasmMode = False
    End If
    
    FileAbsolutePath = File1.Path & "\" & File1.FileName
    
    If disasmMode = True Then
       MsgBox "Selected " & SF2EDITCONF_FILENAME & " : Disasm Mode !", vbOKOnly
       Call LoadDisasm(FileAbsolutePath)
    Else
       'MsgBox "Romfile selected : Classic Mode", vbOKOnly
       Call LoadRom(FileAbsolutePath)
    End If
        
    Unload Me

End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
End Sub

Private Sub File1_DblClick()
 Call cmdOpen_Click
End Sub

Private Sub Form_Load()
 On Error GoTo DP
 
 Dir1.Path = DefaultPath
 
 File1.Path = Dir1.Path
 
 Exit Sub
 
DP:
 
 Dir1.Path = App.Path
 
 File1.Path = Dir1.Path
  
End Sub
