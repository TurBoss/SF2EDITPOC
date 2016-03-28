VERSION 5.00
Begin VB.Form CharacterData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Start Data"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   7350
   Begin VB.Frame Frame3 
      Caption         =   "Inventory"
      Height          =   2655
      Left            =   3840
      TabIndex        =   14
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox cboItems 
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
         TabIndex        =   5
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox chkEquipped 
         Alignment       =   1  'Right Justify
         Caption         =   "Equipped"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   2775
      End
      Begin VB.ListBox lstItems 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   3615
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
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   450
         Width           =   2415
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtNameData 
         Height          =   405
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Class"
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
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Level"
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
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
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
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select"
      Height          =   735
      Left            =   120
      TabIndex        =   9
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
Attribute VB_Name = "CharacterData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Address As Long
Private NameAddress As Long
Private Loading As Boolean

Private Sub cboClass_Click()
  RomDump(Address) = cboClass.ItemData(cboClass.ListIndex)
End Sub

Private Sub cboItems_Click()
 
 txtNameData.Visible = False
 
 If Loading = True Or lstItems.ListIndex = -1 Then
  Exit Sub
 End If

  If chkEquipped.Value = vbChecked Then
  
  lstItems.List(lstItems.ListIndex) = cboItems.Text & " (Equipped)"
  lstItems.ItemData(lstItems.ListIndex) = cboItems.ItemData(cboItems.ListIndex) + &H80
  
  RomDump(Address + 2 + lstItems.ListIndex) = cboItems.ItemData(cboItems.ListIndex) + &H80
 
 Else

  lstItems.List(lstItems.ListIndex) = cboItems.Text
  lstItems.ItemData(lstItems.ListIndex) = cboItems.ItemData(cboItems.ListIndex)
  
  RomDump(Address + 2 + lstItems.ListIndex) = cboItems.ItemData(cboItems.ListIndex)
  
 End If
 
End Sub

Private Sub cboName_Click()
  
  Dim Index As Long
  Dim Leng As Byte
  
  txtNameData.Visible = False
  EnableTextboxes True, Me
   
  Loading = True
    
  Address = GetAddress(cboName.ListIndex)
''  NameAddress = GetNameAddress(cboName.ListIndex)
''
''
''  Leng = Len(cboName.Text)
''  txtName.MaxLength = Leng
''  txtNameData.Text = ""
''
''  For Index = 0 To Leng - 1
''   If Index <> Leng - 1 Then
''   txtNameData.Text = txtNameData.Text & RomDump(NameAddress + Index) & vbNewLine
''   Else
''   txtNameData.Text = txtNameData.Text & RomDump(NameAddress + Index)
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
  
  
  
  
  lstItems.Clear
  cboItems.ListIndex = -1
  chkEquipped.Value = vbUnchecked
    
  SetCombo cboClass, CLng(RomDump(Address))
  
  txtLevel.Text = RomDump(Address + 1)
  
  For Index = 0 To 3
  
   If RomDump(Address + 2 + Index) > &H7F Then
     lstItems.AddItem ItemName(RomDump(Address + 2 + Index) - &H80) & " (Equipped)"
     lstItems.ItemData(Index) = RomDump(Address + 2 + Index)
   Else
     lstItems.AddItem ItemName(RomDump(Address + 2 + Index))
     lstItems.ItemData(Index) = RomDump(Address + 2 + Index)
   End If
   
  Next
  
  Loading = False
  
End Sub


Private Sub chkEquipped_Click()

 txtNameData.Visible = False
 
 If Loading = True Or lstItems.ListIndex = -1 Then
  Exit Sub
 End If
 
 If cboItems.ItemData(cboItems.ListIndex) + &H80 >= 254 Then
  chkEquipped.Value = vbUnchecked
  Exit Sub
 End If
 
 Call cboItems_Click
End Sub

Private Sub cmdRename_Click()
 Dim Rename As New NameSpellHeroMonster
 Dim NameLocation As Integer

 Rename.Show
 
 NameLocation = 44 + CInt(cboName.ListIndex)
   
 Rename.cboName.ListIndex = NameLocation
 
End Sub

Private Sub Form_Load()

 Dim Index As Long
 
 EnableTextboxes False, Me
  
 For Index = 0 To UBound(ItemName())
   cboItems.AddItem ItemName(Index)
   cboItems.ItemData(Index) = ItemCode(Index)
 Next
 
 For Index = 0 To UBound(PersonName())
   cboName.AddItem PersonName(Index)
   cboName.ItemData(Index) = PersonCode(Index)
 Next
 
 For Index = 0 To UBound(ClassName())
   cboClass.AddItem ClassName(Index)
   cboClass.ItemData(Index) = Index
 Next
End Sub

Private Sub txtLevel_Change()
 Dim Result As Byte
 
 If ValidateInput(txtLevel, False, True) = True Then
  Result = txtLevel.Text
  RomDump(Address + 1) = Result
 End If
End Sub

Private Sub txtNameData_KeyDown(KeyCode As Integer, Shift As Integer)
 
 DoEvents
 
 If KeyCode = 13 Then
  txtNameData.Text = Left(txtNameData.Text, txtNameData.SelStart - 2) & Right(txtNameData.Text, Len(txtNameData.Text) - txtNameData.SelStart)
 End If
End Sub

Private Function GetAddress(CharacterNum As Byte) As Long

' Dim Index As Long
' Dim Count As Byte
'
' Index = pStats(0) '&H1EE2F0
'
' Do While Count < CByte(GuyNumber)
'
'  Index = Index + 1
'
'  If RomDump(Index) = &HFF Or RomDump(Index) = &HFE Then
'    Count = Count + 1
'    Index = Index + 1
'  End If
'
' Loop
'
' GetAddress = Index + CharacterNum * 6 + 1
 
 GetAddress = pJoinData + CharacterNum * 6 '+ 1
End Function

Private Sub lstItems_Click()
  
  txtNameData.Visible = False
  
  Loading = True
  
  If lstItems.ItemData(lstItems.ListIndex) > &H7F Then
   chkEquipped.Value = vbChecked
   SetCombo cboItems, CLng(lstItems.ItemData(lstItems.ListIndex) - &H80)
  Else
   chkEquipped.Value = vbUnchecked
   SetCombo cboItems, CLng(lstItems.ItemData(lstItems.ListIndex))
  End If
  
  Loading = False
  
End Sub

Private Sub txtLevel_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtLevel_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtLevel, False) = True Then
  Result = txtLevel.Text
  RomDump(Address + 1) = Result
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
  
  NameArray = Split(txtNameData.Text, vbNewLine)

  For Index = 0 To UBound(NameArray())
   Result = Result & Chr(CLng(CByte(NameArray(Index))))
  Next Index
 
  For Index = 0 To UBound(NameArray())
  
   RomDump(NameAddress + Index) = CByte(NameArray(Index))
   
  Next Index
  
  
  Exit Sub
  
Skipo:
End Sub

Private Function GetNameAddress(ByVal iName As Long) As Long

 Dim Index As Long
 Dim Count As Long
 
 Count = ALLYNAMES_ORIGINAL_OFFSET

 Do While Index < iName
  Count = Count + RomDump(Count) + 1
  
  Index = Index + 1
 Loop
 
 GetNameAddress = Count + 1
 
End Function
