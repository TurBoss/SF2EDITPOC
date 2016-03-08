VERSION 5.00
Begin VB.Form CharacterStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Stats"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   7560
   Tag             =   "1"
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   495
      Left            =   2760
      TabIndex        =   42
      Top             =   1120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Hide Spell List"
      Height          =   495
      Left            =   2790
      TabIndex        =   2
      Top             =   1120
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Caption         =   "Battle Sprite"
      Height          =   855
      Left            =   120
      TabIndex        =   39
      Top             =   6360
      Width           =   4215
      Begin VB.TextBox txtBattleModel 
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtBattlePalette 
         Height          =   285
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Palette"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   40
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4455
      Left            =   120
      TabIndex        =   30
      Top             =   1800
      Width           =   4215
      Begin VB.TextBox txtAtk 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtMpProjected 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtMp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cboMpGrowth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Characters.frx":0000
         Left            =   2760
         List            =   "Characters.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtDefProjected 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   13
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtAtkProjected 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   10
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtAgi 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   15
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox txtDef 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   12
         Top             =   3000
         Width           =   615
      End
      Begin VB.ComboBox cboAgiGrowth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Characters.frx":0004
         Left            =   2760
         List            =   "Characters.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3840
         Width           =   1215
      End
      Begin VB.ComboBox cboDefGrowth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Characters.frx":0008
         Left            =   2760
         List            =   "Characters.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3000
         Width           =   1215
      End
      Begin VB.ComboBox cboAtkGrowth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Characters.frx":000C
         Left            =   2760
         List            =   "Characters.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox cboHpGrowth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Characters.frx":0010
         Left            =   2760
         List            =   "Characters.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtAgiProjected 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   16
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox txtHpProjected 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtHp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "MP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "AGI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "DEF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "ATK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "HP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Growth"
         Height          =   255
         Left            =   3045
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lv.30 Projected"
         Height          =   195
         Left            =   1665
         TabIndex        =   32
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Base"
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "For Class..."
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   960
      Width           =   2535
      Begin VB.ComboBox cboClass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select"
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox cboName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Characters.frx":0014
         Left            =   120
         List            =   "Characters.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spells"
      Height          =   7095
      Left            =   4440
      TabIndex        =   26
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox chkUsePrevious 
         Alignment       =   1  'Right Justify
         Caption         =   "Use Base Class List"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   5520
         Width           =   2535
      End
      Begin VB.CommandButton cmdRemoveSpell 
         Caption         =   "Remove Spell Slot"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   6480
         Width           =   2535
      End
      Begin VB.CommandButton cmdAddSpell 
         Caption         =   "Add Spell Slot"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   6000
         Width           =   2535
      End
      Begin VB.TextBox txtSpellLevel 
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Top             =   5040
         Width           =   855
      End
      Begin VB.ComboBox cboSpell 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   4440
         Width           =   2535
      End
      Begin VB.ListBox lstSpells 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "@ Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   5040
         Width           =   975
      End
   End
End
Attribute VB_Name = "CharacterStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Address As Long
Private CharacterIndex As Long
Private Loading As Boolean

Private SpriteAddress As Long


Private OffsetAddress As Long

Private Sub cboClass_Click()
  RomDump(Address) = cboClass.ItemData(cboClass.ListIndex)
  RomDump(SpriteAddress) = cboClass.ItemData(cboClass.ListIndex)
End Sub

Private Sub cboDefGrowth_Click()
  RomDump(Address + 10) = cboDefGrowth.ItemData(cboDefGrowth.ListIndex)
End Sub

Private Sub cboHpGrowth_Click()
  RomDump(Address + 1) = cboHpGrowth.ItemData(cboHpGrowth.ListIndex)
End Sub

Private Sub cboMpGrowth_Click()
  RomDump(Address + 4) = cboMpGrowth.ItemData(cboMpGrowth.ListIndex)
End Sub

Private Sub cboAtkGrowth_Click()
  RomDump(Address + 7) = cboAtkGrowth.ItemData(cboAtkGrowth.ListIndex)
End Sub

Private Sub cboAgiGrowth_Click()
  RomDump(Address + 13) = cboAgiGrowth.ItemData(cboAgiGrowth.ListIndex)
End Sub

Private Sub cboName_Click()
  Dim ar As Long
   ar = GetOffsetAddress(cboName.ListIndex)
     
  Dim Index As Long
  Dim SubIndex As Long
  
  Loading = True
    
  EnableTextboxes True, Me
   
   
  Call SpriteUpdate
  txtBattleModel.Text = RomDump(SpriteAddress + 1)
  txtBattlePalette.Text = RomDump(SpriteAddress + 2)
  
   
  lstSpells.Clear
  
  cboSpell.ListIndex = -1
  txtSpellLevel.Text = ""
  chkUsePrevious.Value = vbUnchecked
  
  If cboName.ListIndex = 0 Then
   chkUsePrevious.Enabled = False
  Else
   chkUsePrevious.Enabled = True
  End If
  
  Address = GetAddress(cboName.ListIndex)
  CharacterIndex = cboName.ListIndex

  SetCombo cboClass, CLng(RomDump(Address))
  
  txtHP.Text = RomDump(Address + 2)
  txtHpProjected.Text = RomDump(Address + 3)
  
  txtMP.Text = RomDump(Address + 5)
  txtMpProjected.Text = RomDump(Address + 6)
  
  txtAtk.Text = RomDump(Address + 8)
  txtAtkProjected.Text = RomDump(Address + 9)
  
  txtDef.Text = RomDump(Address + 11)
  txtDefProjected.Text = RomDump(Address + 12)
  
  txtAGI.Text = RomDump(Address + 14)
  txtAgiProjected.Text = RomDump(Address + 15)
    
  cboHpGrowth.ListIndex = RomDump(Address + 1)
  cboMpGrowth.ListIndex = RomDump(Address + 4)
  cboAtkGrowth.ListIndex = RomDump(Address + 7)
  cboDefGrowth.ListIndex = RomDump(Address + 10)
  cboAgiGrowth.ListIndex = RomDump(Address + 13)
  
  Index = Address + 16
  
  Do While RomDump(Index) <> &HFF And RomDump(Index) <> &HFE
      
    SubIndex = 0
    
    Do While SpellCode(SubIndex) <> RomDump(Index + 1) And SubIndex < UBound(SpellCode())
     SubIndex = SubIndex + 1
    Loop
    
    lstSpells.AddItem SpellName(SubIndex)
    lstSpells.ItemData(lstSpells.ListCount - 1) = SpellCode(SubIndex)
  
    Index = Index + 2
  Loop
 
  chkUsePrevious.Tag = Index
  
  If RomDump(Index) = 254 Then
   chkUsePrevious.Value = vbChecked
  End If
  
  Loading = False
End Sub

Private Sub cbospell_Click()
    
  If lstSpells.ListIndex = -1 Or Loading = True Then
   Exit Sub
  End If
  
  lstSpells.List(lstSpells.ListIndex) = cboSpell.Text
  lstSpells.ItemData(lstSpells.ListIndex) = cboSpell.ItemData(cboSpell.ListIndex)
  
  RomDump(Address + 17 + lstSpells.ListIndex * 2) = cboSpell.ItemData(cboSpell.ListIndex)
  
End Sub

Private Sub chkUsePrevious_Click()
  If Loading = True Then
   Exit Sub
  End If
  
  If chkUsePrevious.Value = vbChecked Then
   RomDump(CLng(chkUsePrevious.Tag)) = 254
  Else
   RomDump(CLng(chkUsePrevious.Tag)) = 255
  End If
End Sub

Private Sub cmdAddSpell_Click()
 
 Dim Index As Long
 Dim Beginning As Long
 Dim SubIndex As Long
 Dim Count As Byte
 Dim Iterations As Long
 
 Index = pStats(0)   ' &H1EE2F0

 If cboName.ListIndex < 0 Then
  Exit Sub
 End If
 
 Beginning = CLng(chkUsePrevious.Tag)
 
' Do While Count <= UBound(CharacterName())

 Do While Count <= GuyNumber - 1

  Index = Index + 1

  If RomDump(Index) = &HFF Or RomDump(Index) = &HFE Then
    Count = Count + 1
    Index = Index + 1
  End If
 
 Loop
 
 'Index = Index + 32 * 6 + 1 + CharacterNum * 5
 
 Index = Index + 32 * 6 + 1 + 32 * 5
 
 If Index - &H1EE930 > 500 And isExpanded = False Then
  MsgBox "This program limits expanding this section beyond 500 times", vbOKOnly
  Exit Sub
 End If
 
 For SubIndex = Index + 2 To Beginning + 1 Step -2
  RomDump(SubIndex) = RomDump(SubIndex - 2)
  RomDump(SubIndex - 1) = RomDump(SubIndex - 3)
 Next SubIndex
  
 RomDump(Beginning) = 0
 RomDump(Beginning + 1) = 0
 
 Call cboName_Click
 
 
 If isExpanded = False Then
 
 'Update 01010101 Pointer
 If RomDump(&H1EE015 + 2) < 254 Then
   RomDump(&H1EE015 + 2) = RomDump(&H1EE015 + 2) + 2
 Else
   If RomDump(&H1EE015 + 2) = 254 Then
     RomDump(&H1EE015 + 2) = 0
   Else
     RomDump(&H1EE015 + 2) = 1
   End If
 
   RomDump(&H1EE015 + 1) = RomDump(&H1EE015 + 1) + 1
 End If
 
 End If
 
 'Update Class Pointer
 If RomDump(&H1EE00D + 2) < 254 Then
   RomDump(&H1EE00D + 2) = RomDump(&H1EE00D + 2) + 2
 Else
   If RomDump(&H1EE00D + 2) = 254 Then
     RomDump(&H1EE00D + 2) = 0
   Else
     RomDump(&H1EE00D + 2) = 1
   End If
 
   RomDump(&H1EE00D + 1) = RomDump(&H1EE00D + 1) + 1
 End If
  
 
 'Update Join Data Pointer
 If RomDump(&H1EE009 + 2) < 254 Then
   RomDump(&H1EE009 + 2) = RomDump(&H1EE009 + 2) + 2
 Else
   If RomDump(&H1EE009 + 2) = 254 Then
     RomDump(&H1EE009 + 2) = 0
   Else
     RomDump(&H1EE009 + 2) = 1
   End If
 
   RomDump(&H1EE009 + 1) = RomDump(&H1EE009 + 1) + 1
 End If
 
 
 'Update the CHARACTER pointer
 
  For Index = cboName.ListIndex To 0 Step -1
   If Index > 0 Then
    If Left(cboName.List(Index), 3) <> Left(cboName.List(Index - 1), 3) Then
     Iterations = Iterations + 1
    End If
   
   End If
 
  Next Index
 
   
'''  For Index = Iterations + 1 To 29
'''
''' If RomDump(&H1EE271 + 2 + 4 * Index) < 254 Then
'''   RomDump(&H1EE271 + 2 + 4 * Index) = RomDump(&H1EE271 + 2 + 4 * Index) + 2
''' Else
'''   If RomDump(&H1EE271 + 2 + 4 * Index) = 254 Then
'''     RomDump(&H1EE271 + 2 + 4 * Index) = 0
'''   Else
'''     RomDump(&H1EE271 + 2 + 4 * Index) = 1
'''   End If
'''
'''   RomDump(&H1EE271 + 1 + 4 * Index) = RomDump(&H1EE271 + 1 + 4 * Index) + 2
''' End If
'''
'''  Next Index
 
 For Index = Iterations + 1 To UBound(GuyBlocks())   '29

 If RomDump(&H1EE271 + 2 + 4 * Index) < 254 Then
   RomDump(&H1EE271 + 2 + 4 * Index) = RomDump(&H1EE271 + 2 + 4 * Index) + 2
 Else
   If RomDump(&H1EE271 + 2 + 4 * Index) = 254 Then
     RomDump(&H1EE271 + 2 + 4 * Index) = 0
   Else
     RomDump(&H1EE271 + 2 + 4 * Index) = 1
   End If

   RomDump(&H1EE271 + 1 + 4 * Index) = RomDump(&H1EE271 + 1 + 4 * Index) + 1 '' 2 isn't the same number as 1 lolololol
 End If

 Next Index

 ReloadPStats
 
 InitializeAddresses
End Sub

Private Sub cmdExport_Click()

 Dim Stringo As String
 Dim Index As Long
 
 Stringo = Stringo & Left(cboName.Text, Len(cboName.Text) - 2)
 Stringo = Stringo & " - " & cboClass.Text
 Stringo = Stringo & vbNewLine & vbNewLine & "HP " & txtHP.Text & " - " & txtHpProjected.Text & " " & cboHpGrowth.Text
 Stringo = Stringo & vbNewLine & "MP " & txtMP.Text & " - " & txtMpProjected.Text & " " & cboMpGrowth.Text
 Stringo = Stringo & vbNewLine & "AT " & txtAtk.Text & " - " & txtAtkProjected.Text & " " & cboAtkGrowth.Text
 Stringo = Stringo & vbNewLine & "DF " & txtDef.Text & " - " & txtDefProjected.Text & " " & cboDefGrowth.Text
 Stringo = Stringo & vbNewLine & "AG " & txtAGI.Text & " - " & txtAgiProjected.Text & " " & cboAgiGrowth.Text
 
 
 Reader.Show
 Reader.Text.Text = Stringo
 
 
End Sub

Private Sub cmdRemoveSpell_Click()

 Dim Index As Long
 Dim Beginning As Long
 Dim SubIndex As Long
 Dim Count As Byte
 Dim Iterations As Long
 
 Index = pStats(0) ' &H1EE2F0

 If cboName.ListIndex < 0 Or lstSpells.ListCount = 0 Then
  Exit Sub
 End If
 
 Beginning = CLng(chkUsePrevious.Tag)
 
 'Do While Count <= UBound(CharacterName())
 Do While Count <= GuyNumber - 1
  
  Index = Index + 1

  If RomDump(Index) = &HFF Or RomDump(Index) = &HFE Then
    Count = Count + 1
    Index = Index + 1
  End If
 
 Loop
 
 'Index = Index + 32 * 6 + 1 + CharacterNum * 5
 
 Index = Index + 32 * 6 + 1 + 32 * 5
 
 '''RomDump(Beginning - 2) = RomDump(Beginning)
 
 For SubIndex = Beginning - 2 To Index Step 2
  RomDump(SubIndex) = RomDump(SubIndex + 2)
  RomDump(SubIndex + 1) = RomDump(SubIndex + 3)
 Next SubIndex
  
 ''RomDump(Beginning) = 0
 ''RomDump(Beginning + 1) = 0
 
 Call cboName_Click
 
 If isExpanded = False Then
 
 'Update 01010101 Pointer
 If RomDump(&H1EE015 + 2) > 1 Then
   RomDump(&H1EE015 + 2) = RomDump(&H1EE015 + 2) - 2
 Else
   If RomDump(&H1EE015 + 2) = 1 Then
     RomDump(&H1EE015 + 2) = 255
   Else
     RomDump(&H1EE015 + 2) = 254
   End If
 
   RomDump(&H1EE015 + 1) = RomDump(&H1EE015 + 1) - 1
 End If
 
 End If
 
 'Update Class Pointer
 If RomDump(&H1EE00D + 2) > 1 Then
   RomDump(&H1EE00D + 2) = RomDump(&H1EE00D + 2) - 2
 Else
   If RomDump(&H1EE00D + 2) = 1 Then
     RomDump(&H1EE00D + 2) = 255
   Else
     RomDump(&H1EE00D + 2) = 254
   End If
 
   RomDump(&H1EE00D + 1) = RomDump(&H1EE00D + 1) - 1
 End If
  
 
 'Update Join Data Pointer
 If RomDump(&H1EE009 + 2) > 1 Then
   RomDump(&H1EE009 + 2) = RomDump(&H1EE009 + 2) - 2
 Else
   If RomDump(&H1EE009 + 2) = 1 Then
     RomDump(&H1EE009 + 2) = 255
   Else
     RomDump(&H1EE009 + 2) = 254
   End If
 
   RomDump(&H1EE009 + 1) = RomDump(&H1EE009 + 1) - 1
 End If
 
 
 'Update the CHARACTER pointer
 
  For Index = cboName.ListIndex To 0 Step -1
   If Index > 0 Then
    If Left(cboName.List(Index), 3) <> Left(cboName.List(Index - 1), 3) Then
     Iterations = Iterations + 1
    End If
   
   End If
 
  Next Index
 
  
''  For Index = Iterations + 1 To 29
''
'' If RomDump(&H1EE271 + 2 + 4 * Index) > 1 Then
''   RomDump(&H1EE271 + 2 + 4 * Index) = RomDump(&H1EE271 + 2 + 4 * Index) - 2
'' Else
''   If RomDump(&H1EE271 + 2 + 4 * Index) = 1 Then
''     RomDump(&H1EE271 + 2 + 4 * Index) = 255
''   Else
''     RomDump(&H1EE271 + 2 + 4 * Index) = 254
''   End If
''
''   RomDump(&H1EE271 + 1 + 4 * Index) = RomDump(&H1EE271 + 1 + 4 * Index) - 2
'' End If
''
''  Next Index
 
 For Index = Iterations + 1 To UBound(GuyBlocks())   '29
 
 If RomDump(&H1EE271 + 2 + 4 * Index) > 1 Then
   RomDump(&H1EE271 + 2 + 4 * Index) = RomDump(&H1EE271 + 2 + 4 * Index) - 2
 Else
   If RomDump(&H1EE271 + 2 + 4 * Index) = 1 Then
     RomDump(&H1EE271 + 2 + 4 * Index) = 255
   Else
     RomDump(&H1EE271 + 2 + 4 * Index) = 254
   End If
 
   RomDump(&H1EE271 + 1 + 4 * Index) = RomDump(&H1EE271 + 1 + 4 * Index) - 1
 End If
  
 Next Index
 
 ReloadPStats
 
 Call InitializeAddresses
End Sub

Private Sub cmdShow_Click()
  
  If Me.Width = 7650 Then
   Me.Width = 4525
   cmdShow.Caption = "Show Spell List"
  Else
   Me.Width = 7650
   cmdShow.Caption = "Hide Spell List"
  End If

End Sub

Private Sub Form_Load()
 
 Dim Index As Long
 Dim SubIndex As Long
 
  EnableTextboxes False, Me
'''' For Index = 0 To 16
''''  lstSpells.AddItem "Bolt " & Index
''''
'''' Next
 
 For Index = 0 To UBound(SpellName())
   cboSpell.AddItem SpellName(Index)
   cboSpell.ItemData(Index) = SpellCode(Index)
 Next
 
 Call LoadNamebox
 
 For Index = 0 To UBound(ClassName())
   cboClass.AddItem ClassName(Index)
   cboClass.ItemData(Index) = Index
 Next
 
 For Index = 0 To UBound(StatGrowthName())
   
   cboHpGrowth.AddItem StatGrowthName(Index)
   cboHpGrowth.ItemData(Index) = StatGrowthCode(Index)
   
   cboMpGrowth.AddItem StatGrowthName(Index)
   cboMpGrowth.ItemData(Index) = StatGrowthCode(Index)
   
   cboAtkGrowth.AddItem StatGrowthName(Index)
   cboAtkGrowth.ItemData(Index) = StatGrowthCode(Index)
   
   cboDefGrowth.AddItem StatGrowthName(Index)
   cboDefGrowth.ItemData(Index) = StatGrowthCode(Index)
   
   cboAgiGrowth.AddItem StatGrowthName(Index)
   cboAgiGrowth.ItemData(Index) = StatGrowthCode(Index)
    
 Next
 
 'Me = &H1EE2F0
End Sub


Private Function GetAddress(CharacterNum As Byte) As Long

 Dim Index As Long
 Dim Count As Byte

 Index = pStats(0)  '&H1EE2F0

 Do While Count <> CharacterNum

  Index = Index + 1

  If RomDump(Index) = &HFF Or RomDump(Index) = &HFE Then
    Count = Count + 1
    Index = Index + 1
  End If

 Loop

 GetAddress = Index
 
End Function

Private Sub Form_Unload(Cancel As Integer)
 Me.Tag = 1
End Sub

Private Sub lstSpells_Click()

 Dim Index As Long
 
 Loading = True
 
 Index = Address + 16
 
 SetCombo cboSpell, CLng(RomDump(Index + 2 * lstSpells.ListIndex + 1))
 txtSpellLevel.Text = RomDump(Index + 2 * lstSpells.ListIndex)

 Loading = False
End Sub

Private Sub txtAGI_Change()
 Dim Result As Byte
 
 If ValidateInput(txtAGI, False, True, 2) = True Then
  Result = txtAGI.Text
  RomDump(Address + 14) = Result
 End If
End Sub

Private Sub txtAgiProjected_Change()
 Dim Result As Byte
 
 If ValidateInput(txtAgiProjected, False, True, 2) = True Then
  Result = txtAgiProjected.Text
  RomDump(Address + 15) = Result
 End If
End Sub

Private Sub txtAgiProjected_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtAgiProjected, False, , 2) = True Then
  Result = txtAgiProjected.Text
  RomDump(Address + 15) = Result
 End If
 
End Sub

Private Sub txtAgi_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtAGI, False, , 2) = True Then
  Result = txtAGI.Text
  RomDump(Address + 14) = Result
 End If
 
End Sub

Private Sub txtATK_Change()
 Dim Result As Byte
 
 If ValidateInput(txtAtk, False, True, 2) = True Then
  Result = txtAtk.Text
  RomDump(Address + 8) = Result
 End If
End Sub

Private Sub txtAtkProjected_Change()
 Dim Result As Byte
 
 If ValidateInput(txtAtkProjected, False, True, 2) = True Then
  Result = txtAtkProjected.Text
  RomDump(Address + 9) = Result
 End If
End Sub

Private Sub txtAtkProjected_LostFocus()

 Dim Result As Byte
 
 If ValidateInput(txtAtkProjected, False, , 2) = True Then
  Result = txtAtkProjected.Text
  RomDump(Address + 9) = Result
 End If
 
End Sub

Private Sub txtAtk_LostFocus()

 Dim Result As Byte
 
 If ValidateInput(txtAtk, False, , 2) = True Then
  Result = txtAtk.Text
  RomDump(Address + 8) = Result
 End If
 
End Sub



Private Sub txtBattleModel_Change()
 Dim Result As Byte
 
 If ValidateInput(txtBattleModel, False, True, 2) = True Then
  Result = txtBattleModel.Text
  RomDump(SpriteAddress + 1) = Result
 End If
End Sub

Private Sub txtBattleModel_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtBattleModel, False, , 2) = True Then
  Result = txtBattleModel.Text
  RomDump(SpriteAddress + 1) = Result
 End If
 
End Sub

Private Sub txtBattlePalette_Change()
 Dim Result As Byte
 
 If ValidateInput(txtBattlePalette, False, True, 2) = True Then
  Result = txtBattlePalette.Text
  RomDump(SpriteAddress + 2) = Result
 End If
End Sub

Private Sub txtBattlePalette_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtBattlePalette, False, , 2) = True Then
  Result = txtBattlePalette.Text
  RomDump(SpriteAddress + 2) = Result
 End If
 
End Sub

Private Sub txtDEF_Change()
 Dim Result As Byte
 
 If ValidateInput(txtDef, False, True, 2) = True Then
  Result = txtDef.Text
  RomDump(Address + 11) = Result
 End If
End Sub

Private Sub txtDefProjected_Change()
 Dim Result As Byte
 
 If ValidateInput(txtDefProjected, False, True, 2) = True Then
  Result = txtDefProjected.Text
  RomDump(Address + 12) = Result
 End If
End Sub

Private Sub txtDefProjected_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtDefProjected, False, , 2) = True Then
  Result = txtDefProjected.Text
  RomDump(Address + 12) = Result
 End If
 
End Sub

Private Sub txtDef_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtDef, False, , 2) = True Then
  Result = txtDef.Text
  RomDump(Address + 11) = Result
 End If
 
End Sub

Private Sub txtHP_Change()
 Dim Result As Byte
 
 If ValidateInput(txtHP, False, True, 2) = True Then
  Result = txtHP.Text
  RomDump(Address + 2) = Result
 End If
End Sub

Private Sub txtHpProjected_Change()
 Dim Result As Byte
 
 If ValidateInput(txtHpProjected, False, True, 2) = True Then
  Result = txtHpProjected.Text
  RomDump(Address + 3) = Result
 End If
End Sub

Private Sub txtHpProjected_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtHpProjected, False, , 2) = True Then
  Result = txtHpProjected.Text
  RomDump(Address + 3) = Result
 End If
 
End Sub

Private Sub txtHp_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtHP, False, , 2) = True Then
  Result = txtHP.Text
  RomDump(Address + 2) = Result
 End If
 
End Sub

Private Sub txtMP_Change()
 Dim Result As Byte
 
 If ValidateInput(txtMP, False, True, 2) = True Then
  Result = txtMP.Text
  RomDump(Address + 5) = Result
 End If
End Sub

Private Sub txtMpProjected_Change()
 Dim Result As Byte
 
 If ValidateInput(txtMpProjected, False, True, 2) = True Then
  Result = txtMpProjected.Text
  RomDump(Address + 6) = Result
 End If
End Sub

Private Sub txtMpProjected_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtMpProjected, False, , 2) = True Then
  Result = txtMpProjected.Text
  RomDump(Address + 6) = Result
 End If
 
End Sub

Private Sub txtMp_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtMP, False, , 2) = True Then
  Result = txtMP.Text
  RomDump(Address + 5) = Result
 End If
 
End Sub

Private Sub txtSpellLevel_Change()
  Dim Result As Byte
  
  DoEvents
  
  If lstSpells.ListIndex = -1 Or Loading = True Then
   Exit Sub
  End If

  If ValidateInput(txtSpellLevel, False, True, 2) = True Then
   Result = txtSpellLevel.Text
   RomDump(Address + 16 + lstSpells.ListIndex * 2) = Result
  End If
End Sub
'''''
'''''Private Sub txtSpellLevel_LostFocus()
'''''
'''''  Dim Result As Byte
'''''
'''''  DoEvents
'''''
'''''  If lstSpells.ListIndex = -1 Or Loading = True Then
'''''   Exit Sub
'''''  End If
'''''
'''''  If ValidateInput(txtSpellLevel, False, , 2) = True Then
'''''   Result = txtSpellLevel.Text
'''''   RomDump(Address + 16 + lstSpells.ListIndex * 2) = Result
'''''  End If
'''''End Sub

Private Function GetOffsetAddress(iIndex As Long) As Long
 Dim Index As Long
 Dim Count As Long
 Dim Iterations As Long
 
 Count = &H1EE271

 For Index = iIndex To 0 Step -1
   If Index > 0 Then
    If Left(cboName.List(Index), 3) <> Left(cboName.List(Index - 1), 3) Then
     Iterations = Iterations + 1
    End If
   
   End If
 
 Next Index

 Index = 0
 
 Do While Index < Iterations
  Count = Count + 4
  
  Index = Index + 1
 Loop
 
 GetOffsetAddress = Count
End Function


Private Sub SpriteUpdate()

 Dim Counts As Long
 Dim Offset As Integer
 Dim Index As Long
 
 Counts = pBattleSprites ' &H1F806
 
 Index = 0
 
 Offset = cboName.ListIndex
 
 Do While Index < Offset
 
  If RomDump(Counts + 3) < 255 Then
    Counts = Counts + 3
    Index = Index + 1
  Else
    Counts = Counts + 3
  End If
 
 Loop
  
 SpriteAddress = Counts
End Sub


Private Sub LoadNamebox()
  
 Dim Index As Long
 Dim SubIndex As Long
 
 For Index = 1 To GuyBlocks(0)
   cboName.AddItem PersonName(0) & " " & Index   'CharacterName(0)
 Next Index
 
 For Index = 1 To UBound(GuyBlocks())  '29
   For SubIndex = GuyBlocks(Index - 1) To GuyBlocks(Index) - 1
    cboName.AddItem PersonName(Index) & " " & SubIndex - GuyBlocks(Index - 1) + 1 ' CharacterName(Index)
   Next SubIndex
 Next

End Sub
