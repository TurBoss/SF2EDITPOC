VERSION 5.00
Begin VB.Form Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   8805
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSortIndex 
      Caption         =   "Sort by Index"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdSortName 
      Caption         =   "Sort by Name"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Equipable Class"
      Height          =   4095
      Left            =   120
      TabIndex        =   56
      Top             =   3360
      Width           =   8535
      Begin VB.CheckBox chkClass 
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
         Index           =   31
         Left            =   6480
         TabIndex        =   45
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   30
         Left            =   6480
         TabIndex        =   44
         Top             =   3135
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   29
         Left            =   6480
         TabIndex        =   43
         Top             =   2670
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   28
         Left            =   6480
         TabIndex        =   42
         Top             =   2205
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   27
         Left            =   6480
         TabIndex        =   41
         Top             =   1740
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   26
         Left            =   6480
         TabIndex        =   40
         Top             =   1290
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   25
         Left            =   6480
         TabIndex        =   39
         Top             =   825
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   24
         Left            =   6480
         TabIndex        =   38
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   23
         Left            =   4400
         TabIndex        =   37
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   22
         Left            =   4400
         TabIndex        =   36
         Top             =   3135
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   21
         Left            =   4400
         TabIndex        =   35
         Top             =   2670
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   20
         Left            =   4400
         TabIndex        =   34
         Top             =   2205
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   19
         Left            =   4400
         TabIndex        =   33
         Top             =   1740
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   18
         Left            =   4400
         TabIndex        =   32
         Top             =   1290
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   17
         Left            =   4400
         TabIndex        =   31
         Top             =   825
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   16
         Left            =   4400
         TabIndex        =   30
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   15
         Left            =   2320
         TabIndex        =   29
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   14
         Left            =   2320
         TabIndex        =   28
         Top             =   3135
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   13
         Left            =   2320
         TabIndex        =   27
         Top             =   2670
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   12
         Left            =   2320
         TabIndex        =   26
         Top             =   2205
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   11
         Left            =   2320
         TabIndex        =   25
         Top             =   1740
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   10
         Left            =   2320
         TabIndex        =   24
         Top             =   1290
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   9
         Left            =   2320
         TabIndex        =   23
         Top             =   825
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   8
         Left            =   2320
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   3132
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   2670
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   2208
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1746
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1284
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   822
         Width           =   1695
      End
      Begin VB.CheckBox chkClass 
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
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   47
      Top             =   960
      Width           =   8535
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtMaxRange 
         Height          =   285
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtMinRange 
         Height          =   285
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox cboSpell 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtAttribute2 
         Height          =   285
         Left            =   4440
         MaxLength       =   3
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtAttribute1 
         Height          =   285
         Left            =   4440
         MaxLength       =   3
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cboAttributeCode2 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cboAttributeCode1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtNameData 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   5400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   55
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   54
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Attribute 2"
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
         TabIndex        =   53
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Attribute 1"
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
         TabIndex        =   52
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Spell"
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
         Left            =   4440
         TabIndex        =   51
         Top             =   1800
         Width           =   855
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
         Left            =   4440
         TabIndex        =   50
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Value (GP)"
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
         TabIndex        =   49
         Top             =   1800
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
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      Height          =   735
      Left            =   120
      TabIndex        =   46
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox cboNameSorted 
         Height          =   315
         ItemData        =   "Item.frx":0000
         Left            =   120
         List            =   "Item.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cboName 
         Height          =   315
         ItemData        =   "Item.frx":0004
         Left            =   120
         List            =   "Item.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ItemIndex As Long
Private Loading As Boolean


Private Sub cboName_Click()
  EnableTextboxes True, Me
   
  Load_Item cboName.List(cboName.ListIndex), cboName.ItemData(cboName.ListIndex)
End Sub

Private Sub cboNameSorted_Click()
 EnableTextboxes True, Me
 
 txtNameData.Visible = False
 
 Load_Item cboNameSorted.List(cboNameSorted.ListIndex), cboNameSorted.ItemData(cboNameSorted.ListIndex)
End Sub




'Private Sub Form_Load()
'


Private Sub chkClass_Click(Index As Integer)
   
  txtNameData.Visible = False
  
  If Loading = False Then
  
  RomDump(ItemIndex) = 0
  RomDump(ItemIndex + 1) = 0
  RomDump(ItemIndex + 2) = 0
  RomDump(ItemIndex + 3) = 0
   
  For Index = 0 To chkClass.UBound '31
  
   If chkClass(Index).Value = vbChecked Then
     RomDump(ItemIndex + ItemEquipBank(Index) - 1) = RomDump(ItemIndex + ItemEquipBank(Index) - 1) + ItemEquipCode(Index)
   End If
   
''   If ItemEquipCode(Index) <= EquipByte(ItemEquipBank(Index) - 1) Then
''    EquipByte(ItemEquipBank(Index) - 1) = EquipByte(ItemEquipBank(Index) - 1) - ItemEquipCode(Index)
''    chkClass(Index).Value = vbChecked
''   Else
''    chkClass(Index).Value = vbUnchecked
''   End If
   
  Next Index
  
  
  End If
  
End Sub

Private Sub cmdRename_Click()
 Dim Rename As New NameItemClass
 Dim NameLocation As Integer

 Rename.Show
 
 NameLocation = CInt(cboName.ListIndex)
   
 Rename.cboName.ListIndex = NameLocation
End Sub

Private Sub cmdSortIndex_Click()
 cboName.Visible = True
 cboNameSorted.Visible = False
 
 cmdRename.Enabled = True
End Sub

Private Sub cmdSortName_Click()
 cboNameSorted.Visible = True
 cboName.Visible = False
 
 cmdRename.Enabled = False
End Sub



Private Sub Form_Load()

 Dim Index As Byte
 Dim SubIndex As Byte
 
 EnableTextboxes False, Me
 
 For Index = 0 To chkClass.UBound  '   31
  chkClass(Index).Caption = ClassName(Index)    ' ItemEquipName(Index)
'  chkClass(Index).Tag = ItemEquipCode(Index)
 Next
  
  
 For Index = 0 To UBound(ItemName())
  cboName.AddItem ItemName(Index)
  cboName.ItemData(Index) = Index
  
  cboNameSorted.AddItem ItemName(Index)
  For SubIndex = 0 To cboNameSorted.ListCount - 1
   If ItemName(Index) = cboNameSorted.List(SubIndex) Then
    cboNameSorted.ItemData(SubIndex) = Index
   End If
  Next SubIndex
 Next
 
 For Index = 0 To UBound(AttributeCode())
  cboAttributeCode1.AddItem AttributeName(Index)
  cboAttributeCode1.ItemData(Index) = AttributeCode(Index)
  
  cboAttributeCode2.AddItem AttributeName(Index)
  cboAttributeCode2.ItemData(Index) = AttributeCode(Index)
  
 Next
 
 For Index = 0 To UBound(TypeName())
  cboType.AddItem TypeName(Index)
  cboType.ItemData(Index) = TypeCode(Index)
 Next
 
 For Index = 0 To UBound(SpellName())
  cboSpell.AddItem SpellName(Index)
  cboSpell.ItemData(Index) = SpellCode(Index)
 Next
 
End Sub

' This feature is more for my own dev use - I know I don't know everything
' as I type this - for the now. This will be faster than looking at things
' manually in ye olde hex editor.

' I didn't expect to use the labels for anything, hence why I left the names
' at default. It's called saving minutes that add up to hours of your life by
' not doing something that doesn't _need_ to be done.

Private Sub Label3_Click()
 Me.Caption = "Items - " & RomDump(ItemIndex + 8)
End Sub

Private Sub Label4_Click()
 Me.Caption = "Items - " & RomDump(ItemIndex + 9)
End Sub

Private Sub Label5_Click()
 Me.Caption = "Items - " & RomDump(ItemIndex + 10)
End Sub

Private Sub Label6_Click()
 Me.Caption = "Items - " & RomDump(ItemIndex + 12)
End Sub

Private Sub Load_Item(ByVal iName As String, ByVal iIndex As Long)

  Dim Index As Long
  Dim Leng As Byte
  Dim ListIndex As Long
  
  Dim EquipByte(0 To 3) As Byte
  
  Loading = True
  
  ItemIndex = cboName.ItemData(iIndex)
  
  Leng = Len(iName)
  txtName.MaxLength = Leng
''  txtNameData.Text = ""
''
''  For Index = 0 To Leng - 1
''   If Index <> Leng - 1 Then
''   txtNameData.Text = txtNameData.Text & RomDump(ItemOffsetName(ItemIndex) + Index) & vbNewLine
''   Else
''   txtNameData.Text = txtNameData.Text & RomDump(ItemOffsetName(ItemIndex) + Index)
''   End If
''  Next
''
''
''  ' This bloc loads in the raw number data used for these names
''  Dim NameArray() As String
''  NameArray = Split(txtNameData.Text, vbNewLine)
''  txtName.Text = ""
''  For Index = 0 To UBound(NameArray())
''   txtName.Text = txtName.Text & Chr(CLng(NameArray(Index)))
''  Next Index
  
  
  ItemIndex = ItemOffset(ItemIndex)
    
  SetCombo cboType, CLng(RomDump(ItemIndex + 8))
  SetCombo cboSpell, CLng(RomDump(ItemIndex + 9))
  
  txtValue.Text = CLng(RomDump(ItemIndex + 6)) * 256 + RomDump(ItemIndex + 7)
    
  txtMinRange.Text = RomDump(ItemIndex + 5)
  txtMaxRange.Text = RomDump(ItemIndex + 4)
   
  SetCombo cboAttributeCode1, CLng(RomDump(ItemIndex + 10))
  SetCombo cboAttributeCode2, CLng(RomDump(ItemIndex + 12))
  
  txtAttribute1.Text = RomDump(ItemIndex + 11)
  txtAttribute2.Text = RomDump(ItemIndex + 13)
  
  ' Equipment table. Hooray.
  
  EquipByte(0) = RomDump(ItemIndex)
  EquipByte(1) = RomDump(ItemIndex + 1)
  EquipByte(2) = RomDump(ItemIndex + 2)
  EquipByte(3) = RomDump(ItemIndex + 3)
  
      
  For Index = 31 To 0 Step -1
  
   If ItemEquipCode(Index) <= EquipByte(ItemEquipBank(Index) - 1) Then
    EquipByte(ItemEquipBank(Index) - 1) = EquipByte(ItemEquipBank(Index) - 1) - ItemEquipCode(Index)
    chkClass(Index).Value = vbChecked
   Else
    chkClass(Index).Value = vbUnchecked
   End If
   
  Next Index

  Loading = False
End Sub

Private Sub cboAttributeCode2_click()
 txtNameData.Visible = False
 
 RomDump(ItemIndex + 12) = cboAttributeCode2.ItemData(cboAttributeCode2.ListIndex)
End Sub

Private Sub cboAttributeCode1_click()
 txtNameData.Visible = False
 
 RomDump(ItemIndex + 10) = cboAttributeCode1.ItemData(cboAttributeCode1.ListIndex)
End Sub

Private Sub cbospell_Click()
 txtNameData.Visible = False
 
 RomDump(ItemIndex + 9) = cboSpell.ItemData(cboSpell.ListIndex)
End Sub

Private Sub cboType_Click()
 txtNameData.Visible = False
 
 RomDump(ItemIndex + 8) = cboType.ItemData(cboType.ListIndex)
End Sub

Private Sub txtAttribute1_Change()
 Dim Result As Byte
 
 If ValidateInput(txtAttribute1, False, True) = True Then
  Result = txtAttribute1.Text
  RomDump(ItemIndex + 11) = Result
 End If
End Sub

Private Sub txtAttribute2_Change()
 Dim Result As Byte
 
 If ValidateInput(txtAttribute2, False, True) = True Then
  Result = txtAttribute2.Text
  RomDump(ItemIndex + 13) = Result
 End If
End Sub

Private Sub txtAttribute2_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtAttribute2, False) = True Then
  Result = txtAttribute2.Text
  RomDump(ItemIndex + 13) = Result
 End If
End Sub

Private Sub txtAttribute1_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtAttribute1, False) = True Then
  Result = txtAttribute1.Text
  RomDump(ItemIndex + 11) = Result
 End If
End Sub

Private Sub txtMaxRange_Change()
 Dim Result As Byte
   
 If ValidateInput(txtMaxRange, False, True) = True Then
  Result = txtMaxRange.Text
  
  If Result <= 3 Then
   RomDump(ItemIndex + 4) = Result
  End If
  
 End If
End Sub

Private Sub txtMaxRange_Click()
txtNameData.Visible = False
End Sub

Private Sub txtMaxRange_LostFocus()
  
 Dim Result As Byte
   
 If ValidateInput(txtMaxRange, False) = True Then
  Result = txtMaxRange.Text
  
  If Result <= 3 Then
   RomDump(ItemIndex + 4) = Result
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
   RomDump(ItemIndex + 5) = Result
  End If
  
 End If
End Sub

Private Sub txtMinRange_Click()
txtNameData.Visible = False
End Sub

Private Sub txtMinRange_LostFocus()
  
 Dim Result As Byte
   
 If ValidateInput(txtMinRange, False) = True Then
  Result = txtMinRange.Text
  
  If Result <= 3 Then
   RomDump(ItemIndex + 5) = Result
  Else
   MsgBox "Invalid input", vbOKOnly
   
   txtMinRange.SetFocus
  End If
  
 End If
  
End Sub

Private Sub txtName_Change()
 
 Dim Index As Long
 Dim Result As String
 
 If Loading = False Then
 
 For Index = 1 To Len(txtName.Text)
  
  Result = Result & ChCode(Mid(txtName, Index, 1)) & vbNewLine
 
 Next Index
 
 Index = 0
 
 Do While Len(txtName.Text) + Index < txtName.MaxLength
 
  Result = Result & "32" & vbNewLine
  
  Index = Index + 1
 Loop
 
 txtNameData.Text = Left(Result, Len(Result) - 2)
 
 End If
End Sub

Private Sub txtName_Click()
 txtNameData.Visible = True
End Sub


Private Sub txtNameData_Change()
  On Error GoTo Skipo
  
  Dim Index As Long
  
  Dim NameArray() As String
  
  Dim Result As String
  
  Dim dIndex As Long
    
  NameArray = Split(txtNameData.Text, vbNewLine)

  For Index = 0 To UBound(NameArray())
   Result = Result & Chr(CLng(CByte(NameArray(Index))))
  Next Index
 
    If cboName.Visible = True Then
     dIndex = cboName.ItemData(cboName.ListIndex)
    Else
     dIndex = cboNameSorted.ItemData(cboNameSorted.ListIndex)
    End If
    
    
  For Index = 0 To UBound(NameArray())
  
   RomDump(ItemOffsetName(dIndex) + Index) = CByte(NameArray(Index))
   
  Next Index
  
  
  Exit Sub
  
Skipo:
End Sub

Private Sub txtNameData_Click()
  txtNameData.Visible = True
End Sub


Private Sub txtNameData_KeyDown(KeyCode As Integer, Shift As Integer)
 
 DoEvents
 
 If KeyCode = 13 Then
  txtNameData.Text = Left(txtNameData.Text, txtNameData.SelStart - 2) & Right(txtNameData.Text, Len(txtNameData.Text) - txtNameData.SelStart)
 End If
End Sub

Private Sub txtNameData_LostFocus()
  
  On Error GoTo Skipo
  
  Dim Index As Long
  
  Dim NameArray() As String
  
  Dim Result As String
  
  Dim dIndex As Long
  
  NameArray = Split(txtNameData.Text, vbNewLine)

  For Index = 0 To UBound(NameArray())
   Result = Result & Chr(CLng(CByte(NameArray(Index))))
  Next Index
  
   
  txtName.Text = Result
  
    If cboName.Visible = True Then
     dIndex = cboName.ItemData(cboName.ListIndex)
    Else
     dIndex = cboNameSorted.ItemData(cboNameSorted.ListIndex)
    End If
  

  For Index = 0 To UBound(NameArray())
  
   RomDump(ItemOffsetName(dIndex) + Index) = CByte(NameArray(Index))
   
  Next Index
   
   
  Exit Sub
  
Skipo:

  MsgBox "Invalid input", vbOKOnly
  
  txtNameData.SetFocus
  
End Sub

Private Sub txtValue_Change()
 Dim Index As Long
 Dim Result As Byte
  
 If ValidateInput(txtValue, True, True) = True Then
  
  Index = CLng(txtValue.Text)
  
  Result = CByte(Fix(Index / 256#))
  
  RomDump(ItemIndex + 6) = Result
  
  Result = CByte(Index - CLng(Result * 256#))
  
  RomDump(ItemIndex + 7) = Result
  
 End If
End Sub

Private Sub txtValue_Click()
txtNameData.Visible = False
End Sub

Private Sub txtValue_LostFocus()
 
 Dim Index As Long
 Dim Result As Byte
  
 If ValidateInput(txtValue, True) = True Then
  
  Index = CLng(txtValue.Text)
  
  Result = CByte(Fix(Index / 256#))
  
  RomDump(ItemIndex + 6) = Result
  
  Result = CByte(Index - CLng(Result * 256#))
  
  RomDump(ItemIndex + 7) = Result
  
 End If
 
 '' txtValue.Text = CLng(RomDump(ItemIndex + 6)) * 256 + RomDump(ItemIndex + 7)
 
End Sub



