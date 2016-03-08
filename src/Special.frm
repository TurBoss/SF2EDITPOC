VERSION 5.00
Begin VB.Form Special 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Special"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   6165
   Begin VB.Frame Frame7 
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   5895
      Begin VB.CheckBox chkTauros 
         Caption         =   "Taurus Can Only Be Damaged By the Achilles Sword"
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
         Left            =   360
         TabIndex        =   26
         Top             =   275
         Width           =   5055
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Stun Paralysis Chance"
      Height          =   855
      Left            =   3480
      TabIndex        =   22
      Top             =   3240
      Width           =   2535
      Begin VB.TextBox txtStun 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "1   /"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   24
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Curse Paralysis Chance"
      Height          =   855
      Left            =   3480
      TabIndex        =   19
      Top             =   2280
      Width           =   2535
      Begin VB.TextBox txtCurse 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   20
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "1   /"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Basic SORC Spell"
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   3255
      Begin VB.ComboBox cboSpells 
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
         TabIndex        =   18
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Attack In Super Mode"
      Height          =   855
      Left            =   3480
      TabIndex        =   14
      Top             =   1160
      Width           =   2535
      Begin VB.TextBox txtSuperAttack 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   960
         MaxLength       =   3
         TabIndex        =   15
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "/   4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Spellpower After Promotion"
      Height          =   855
      Left            =   3480
      TabIndex        =   11
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtSpellpower 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   960
         MaxLength       =   3
         TabIndex        =   13
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "/   4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
   End
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
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kiwi's Atomic Fire Breath"
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtLevel4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtLevel3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtLevel2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtBreathChance 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   1
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Level 4        @  Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2430
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Level 3        @  Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1950
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Level 2        @  Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1470
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Chance to use"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   980
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
   End
End
Attribute VB_Name = "Special"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Loading As Boolean

Private Sub cboName_Click()
  If Loading = False Then
    RomDump(147681) = cboName.ListIndex
  End If
End Sub


Private Sub chkTauros_Click()
 
 
 If chkTauros.Value = vbChecked Then: RomDump(40724) = 102
 If chkTauros.Value = vbUnchecked Then: RomDump(40724) = 96
 
End Sub

Private Sub Form_Load()
 
 Loading = True
 
 EnableTextboxes False, Me
 
 Dim Index As Long

 For Index = 0 To UBound(mSpellName())
  
  cboName.AddItem mSpellName(Index)
  cboName.ItemData(Index) = Index
 Next Index
 
 txtBreathChance.Text = RomDump(147623)
 txtLevel2.Text = RomDump(147655)
 txtLevel3.Text = RomDump(147663)
 txtLevel4.Text = RomDump(147671)
 SetCombo cboName, CLng(RomDump(147681))
  
 For Index = 0 To UBound(SpellName())
   cboSpells.AddItem SpellName(Index)
   cboSpells.ItemData(Index) = SpellCode(Index)
 Next
 
 SetCombo cboSpells, CLng(RomDump(135501))
 
 txtSpellpower.Text = RomDump(47989)
 txtSuperAttack.Text = RomDump(1775083)
 txtCurse.Text = RomDump(40895)
 txtStun.Text = RomDump(40921)
  
 EnableTextboxes True, Me
 
 
 If RomDump(40724) = 102 Then: chkTauros.Value = vbChecked
 
 Loading = False
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtBreathChance_Change()
 Dim Result As Byte
 
 If ValidateInput(txtBreathChance, False, True) = True Then
  Result = txtBreathChance.Text
  RomDump(147623) = Result
 End If
End Sub

Private Sub txtBreathChance_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtBreathChance, False, False) = True Then
  Result = txtBreathChance.Text
  RomDump(147623) = Result
 End If
End Sub

Private Sub txtLevel2_Change()
 Dim Result As Byte
 
 If ValidateInput(txtLevel2, False, True) = True Then
  Result = txtLevel2.Text
  RomDump(147655) = Result
 End If
End Sub

Private Sub txtLevel2_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtLevel2, False, True) = True Then
  Result = txtLevel2.Text
  RomDump(147655) = Result
 End If
End Sub

Private Sub txtLevel3_Change()
 Dim Result As Byte
 
 If ValidateInput(txtLevel3, False, True) = True Then
  Result = txtLevel3.Text
  RomDump(147663) = Result
 End If
End Sub

Private Sub txtLevel3_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtLevel3, False, True) = True Then
  Result = txtLevel3.Text
  RomDump(147663) = Result
 End If
End Sub

Private Sub txtLevel4_Change()
 Dim Result As Byte
 
 If ValidateInput(txtLevel4, False, True) = True Then
  Result = txtLevel4.Text
  RomDump(147671) = Result
 End If
End Sub

Private Sub txtLevel4_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtLevel4, False, True) = True Then
  Result = txtLevel4.Text
  RomDump(147671) = Result
 End If
End Sub

Private Sub txtSpellpower_Change()
 Dim Result As Byte
 
 If ValidateInput(txtSpellpower, False, True) = True Then
  Result = txtSpellpower.Text
  RomDump(47989) = Result
 End If
End Sub

Private Sub txtSpellpower_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtSpellpower, False, True) = True Then
  Result = txtSpellpower.Text
  RomDump(47989) = Result
 End If
End Sub

Private Sub txtSuperAttack_Change()
 Dim Result As Byte
 
 If ValidateInput(txtSuperAttack, False, True) = True Then
  Result = txtSpellpower.Text
  RomDump(1775083) = Result
 End If
End Sub

Private Sub txtSuperAttack_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtSuperAttack, False, True) = True Then
  Result = txtSpellpower.Text
  RomDump(1775083) = Result
 End If
End Sub

Private Sub cboSpells_Click()
  If Loading = True Then
   Exit Sub
  End If
  
  RomDump(135501) = cboSpells.ItemData(cboSpells.ListIndex)
  
End Sub

Private Sub txtCurse_Change()
 Dim Result As Byte
 
 If ValidateInput(txtCurse, False, True) = True Then
  Result = txtCurse.Text
  RomDump(40895) = Result
 End If
End Sub

Private Sub txtCurse_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtCurse, False, True) = True Then
  Result = txtCurse.Text
  RomDump(40895) = Result
 End If
End Sub

Private Sub txtStun_Change()
 Dim Result As Byte
 
 If ValidateInput(txtStun, False, True) = True Then
  Result = txtStun.Text
  RomDump(40921) = Result
 End If
End Sub

Private Sub txtStun_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtStun, False, True) = True Then
  Result = txtStun.Text
  RomDump(40921) = Result
 End If
End Sub



