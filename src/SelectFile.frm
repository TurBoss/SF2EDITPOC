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
      Pattern         =   "*.smd;*.bin;sf2editconf.txt"
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
 Dim Freedfile As Long
 
 On Error GoTo Ouch
 
 Freedfile = FreeFile()
 
 If File1.FileName = "sf2editconf.txt" Then
    MsgBox "Selected sf2editconf.txt : applying Disasm Mode !", vbOKOnly
    disasmMode = True
 Else
    'MsgBox "sf2editconf.txt not selected : Rom Mode", vbOKOnly
    disasmMode = False
 End If
 
 Open File1.Path & "\" & File1.FileName For Binary As #Freedfile
  
  If UCase(Right(File1.FileName, 3)) <> "SMD" Then  'And Right(File1.FileName, 3) <> "SMD" Then
   ReDim RomDump(LOF(Freedfile) - 1)
  Else
   ReDim RomDump(&H1FFFFF)
  End If
  
  Get #Freedfile, , RomDump

 Close #Freedfile
 
 If UCase(Right(File1.FileName, 3)) <> "SMD" Then
  Call InitializeAddresses
 End If
 
 ' Do stuff we couldn't before load
 CalculateStoreSpots
 
 
 RomPath = File1.Path & "\" & File1.FileName
 
 Main.mnuEdit.Enabled = True
 Main.mnuConvert.Enabled = True
 Main.mnuMisc.Enabled = True
 Main.mnuEditNames.Enabled = True
 
 
 Dim Index As Long
 Dim Count As Long
 Dim SubIndex As Long
 
 Index = pItemNames    '&H1796E
 Count = 0
 
 Do While Count <= UBound(mItemName())
 
  mItemNameLength(Count) = RomDump(Index)
    
  Index = Index + 1
    
  mItemName(Count) = ""
    
  For SubIndex = 0 To mItemNameLength(Count) - 1
   mItemName(Count) = mItemName(Count) & Chr(RomDump(Index + SubIndex))
  Next SubIndex
 
  Index = Index + mItemNameLength(Count)
  
  If Count = 127 Then
   Index = Index + 1
  End If
  
  Count = Count + 1
 Loop
 
'' The next name set

 Index = pSpellNames '63940

 Count = 0
 
 Do While Count <= UBound(mSpellName())
 
  mSpellNameLength(Count) = RomDump(Index)
    
  Index = Index + 1
    
  mSpellName(Count) = ""
  
  For SubIndex = 0 To mSpellNameLength(Count) - 1
   mSpellName(Count) = mSpellName(Count) & Chr(RomDump(Index + SubIndex))
  Next SubIndex
 
  Index = Index + mSpellNameLength(Count)
 
  Count = Count + 1
 Loop
 
''' If UCase(Right(File1.FileName, 3)) <> "SMD" Then 'And Right(File1.FileName, 3) <> "SMD" Then
'''  GuyNumber = CountGuys()
''' End If
    
    
 Call LoadRomNames
  
  
 Unload Me
 
 Exit Sub
 
Ouch:
 
 Close Freedfile
 
 MsgBox "The file you selected is incompatible with this program.", vbOKOnly
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
