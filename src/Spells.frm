VERSION 5.00
Begin VB.Form Spells 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spells"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   6525
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   3840
      TabIndex        =   19
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtEffect 
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   5
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Effect"
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
         TabIndex        =   20
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   6255
      Begin VB.TextBox txtAnimation 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtAoE 
         Height          =   285
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtMaxRange 
         Height          =   285
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtMinRange 
         Height          =   285
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtMPCost 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "AoE"
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
         Left            =   4080
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Max Range"
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
         Left            =   4080
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Min Range"
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
         Left            =   4080
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Animation"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Type"
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
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "MP Cost"
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
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
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
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton cmdRename 
         Caption         =   "Rename"
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   1215
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Spells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Address

Private Sub cboName_Click()
    
  '' Me.Caption = cboName.ListIndex
  
  EnableTextboxes True, Me
   
  Address = cboName.ItemData(cboName.ListIndex)

  ''' Me.Caption = RomDump(Address + 2) & " - " & RomDump(Address + 3)
  
  txtMPCost.Text = RomDump(Address + 1)
  txtMaxRange.Text = RomDump(Address + 4)
  txtMinRange.Text = RomDump(Address + 5)
  txtAoE.Text = RomDump(Address + 6)
  txtEffect.Text = RomDump(Address + 7)
  txtAnimation.Text = RomDump(Address + 2)
  
  SetCombo cboType, CLng(RomDump(Address + 3))
End Sub

Private Sub cboType_Click()
   RomDump(Address + 3) = cboType.ItemData(cboType.ListIndex)
End Sub

Private Sub cmdRename_Click()
 Dim Rename As New NameSpellHeroMonster
 Dim NameLocation As Integer

 Rename.Show
 
 NameLocation = CInt(cboName.ListIndex)
 
 
 
 If cboName.ListIndex < 14 Then
  NameLocation = Fix(NameLocation / 4)
 Else
  NameLocation = 3
 End If
 
 If cboName.ListIndex > 13 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 15 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 16 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 17 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 19 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 21 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 22 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 23 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 27 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 31 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 35 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 39 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 40 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 43 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 46 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 49 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 51 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 52 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 53 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 54 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 55 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 56 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 57 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 58 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 61 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 64 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 66 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 68 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 70 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 72 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 73 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 75 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 76 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 77 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 78 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 79 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 80 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 82 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 86 Then: NameLocation = NameLocation + 1
 If cboName.ListIndex > 87 Then: NameLocation = NameLocation + 1
 
 Rename.cboName.ListIndex = NameLocation
 

End Sub

Private Sub Form_Load()
 
 Dim Index As Long
 Dim SubIndex As Byte
 Dim SpellIndex As Byte
 
 EnableTextboxes False, Me
  
 Index = pSpells  '&H176A6
 
 Do
    
   SpellIndex = RomDump(Index)
   
   SubIndex = 0
   
   Do While SpellCode(SubIndex) <> SpellIndex
    SubIndex = SubIndex + 1
   Loop
 
   cboName.AddItem SpellName(SubIndex)
   cboName.ItemData(cboName.ListCount - 1) = Index
   
   Index = Index + 8
 Loop While Index <= pSpells + 704            '&H17966
 
 For Index = 0 To UBound(SpellTypeName())
  cboType.AddItem SpellTypeName(Index)
  cboType.ItemData(Index) = SpellTypeCode(Index)
 Next
End Sub


Private Sub txtAnimation_Change()

 Dim Result As Byte
 
 If ValidateInput(txtAnimation, False, True) = True Then
  Result = txtAnimation.Text
  RomDump(Address + 2) = Result
 End If
 
End Sub

Private Sub txtAnimation_LostFocus()

 Dim Result As Byte
 
 If ValidateInput(txtAnimation, False) = True Then
  Result = txtAnimation.Text
  RomDump(Address + 2) = Result
 End If
 
End Sub

Private Sub txtAoE_Change()
  
  Dim Result As Byte
   
 If ValidateInput(txtAoE, False, True) = True Then
  Result = txtAoE.Text
  
  If Result <= 3 Then
   RomDump(Address + 6) = Result
  Else
   MsgBox "Invalid input", vbOKOnly
   
   txtAoE.SetFocus
  End If
  
 End If
 
End Sub

Private Sub txtAoE_LostFocus()

  Dim Result As Byte
   
 If ValidateInput(txtAoE, False) = True Then
  Result = txtAoE.Text
  
  If Result <= 3 Then
   RomDump(Address + 6) = Result
  Else
   MsgBox "Invalid input", vbOKOnly
   
   txtAoE.SetFocus
  End If
  
 End If
 
End Sub

Private Sub txtEffect_Change()

 Dim Result As Byte
 
 If ValidateInput(txtEffect, False, True) = True Then
  Result = txtEffect.Text
  RomDump(Address + 7) = Result
 End If
 
End Sub

Private Sub txtEffect_LostFocus()

 Dim Result As Byte
 
 If ValidateInput(txtEffect, False) = True Then
  Result = txtEffect.Text
  RomDump(Address + 7) = Result
 End If

End Sub

Private Sub txtMaxRange_Change()

 Dim Result As Byte
   
 If ValidateInput(txtMaxRange, False, True) = True Then
  Result = txtMaxRange.Text
  
  If Result <= 3 Then
   RomDump(Address + 4) = Result
  End If
  
 End If
 
End Sub

Private Sub txtMaxRange_LostFocus()
  
 Dim Result As Byte
   
 If ValidateInput(txtMaxRange, False) = True Then
  Result = txtMaxRange.Text
  
  If Result <= 3 Then
   RomDump(Address + 4) = Result
  Else
   MsgBox "Invalid input", vbOKOnly
   
   txtMaxRange.SetFocus
  End If
  
 End If
End Sub

Private Sub txtMinRange_Change()

 Dim Result As Byte
   
 If ValidateInput(txtMinRange, False, True) = True Then
  Result = txtMinRange.Text
  
  If Result <= 3 Then
   RomDump(Address + 5) = Result
  End If
  
 End If
 
End Sub

Private Sub txtMinRange_LostFocus()

 Dim Result As Byte
   
 If ValidateInput(txtMinRange, False) = True Then
  Result = txtMinRange.Text
  
  If Result <= 3 Then
   RomDump(Address + 5) = Result
  Else
   MsgBox "Invalid input", vbOKOnly
   
   txtMinRange.SetFocus
  End If
  
 End If
 
End Sub

Private Sub txtMPCost_Change()
 
 Dim Result As Byte
 
 If ValidateInput(txtMPCost, False, True) = True Then
  Result = txtMPCost.Text
  RomDump(Address + 1) = Result
 End If
 
End Sub

Private Sub txtMPCost_LostFocus()

 Dim Result As Byte
 
 If ValidateInput(txtMPCost, False) = True Then
  Result = txtMPCost.Text
  RomDump(Address + 1) = Result
 End If
 
End Sub
