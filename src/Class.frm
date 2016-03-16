VERSION 5.00
Begin VB.Form Class 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Class"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   7560
   Begin VB.Frame Frame7 
      Caption         =   "Critical Attack Chance"
      Height          =   2295
      Left            =   120
      TabIndex        =   35
      Top             =   3480
      Width           =   7335
      Begin VB.OptionButton optCritical 
         Caption         =   "Muddle"
         Enabled         =   0   'False
         Height          =   495
         Index           =   12
         Left            =   2640
         TabIndex        =   51
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Slow"
         Enabled         =   0   'False
         Height          =   495
         Index           =   13
         Left            =   3920
         TabIndex        =   50
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Silence"
         Enabled         =   0   'False
         Height          =   495
         Index           =   15
         Left            =   6360
         TabIndex        =   49
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "MP Drain"
         Enabled         =   0   'False
         Height          =   495
         Index           =   14
         Left            =   5080
         TabIndex        =   48
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "None"
         Enabled         =   0   'False
         Height          =   495
         Index           =   8
         Left            =   2640
         TabIndex        =   47
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Poison"
         Enabled         =   0   'False
         Height          =   495
         Index           =   9
         Left            =   3920
         TabIndex        =   46
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Stun"
         Enabled         =   0   'False
         Height          =   495
         Index           =   11
         Left            =   6360
         TabIndex        =   45
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "Sleep"
         Enabled         =   0   'False
         Height          =   495
         Index           =   10
         Left            =   5080
         TabIndex        =   44
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 4"
         Enabled         =   0   'False
         Height          =   495
         Index           =   6
         Left            =   2640
         TabIndex        =   43
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 8"
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   3920
         TabIndex        =   42
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 32"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   6360
         TabIndex        =   41
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 16"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   5080
         TabIndex        =   40
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 4"
         Enabled         =   0   'False
         Height          =   495
         Index           =   7
         Left            =   2640
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 8"
         Enabled         =   0   'False
         Height          =   495
         Index           =   5
         Left            =   3920
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 32"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   6360
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCritical 
         Caption         =   "1 / 16"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   5080
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "150% Damage"
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
         TabIndex        =   53
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "125% Damage"
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
         TabIndex        =   52
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Double Attack Chance"
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   5880
      Width           =   3615
      Begin VB.OptionButton optDouble 
         Caption         =   "1 / 16"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   1770
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optDouble 
         Caption         =   "1 / 32"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optDouble 
         Caption         =   "1 / 8"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   960
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optDouble 
         Caption         =   "1 / 4"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Counter Attack Chance"
      Height          =   855
      Left            =   3840
      TabIndex        =   25
      Top             =   5880
      Width           =   3615
      Begin VB.OptionButton optCounter 
         Caption         =   "1 / 4"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCounter 
         Caption         =   "1 / 8"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   960
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optCounter 
         Caption         =   "1 / 32"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCounter 
         Caption         =   "1 / 16"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   1770
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   3840
      TabIndex        =   23
      Top             =   2640
      Width           =   3615
      Begin VB.TextBox txtOtherResistance 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Other Resistance"
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
         TabIndex        =   24
         Top             =   270
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Resistance - Select Both For Weakness"
      Height          =   2415
      Left            =   3840
      TabIndex        =   21
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox chk25Wind 
         Caption         =   "25% Wind"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Tag             =   "1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chk50Wind 
         Caption         =   "50% Wind"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Tag             =   "2"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chk25Lightning 
         Caption         =   "25% Lightning"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "4"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chk50Lightning 
         Caption         =   "50% Lightning"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Tag             =   "8"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chk25Ice 
         Caption         =   "25% Ice"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Tag             =   "16"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox chk50Ice 
         Caption         =   "50% Ice"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Tag             =   "32"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox chk25Fire 
         Caption         =   "25% Fire"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Tag             =   "64"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox chk50Fire 
         Caption         =   "50% Fire"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Tag             =   "128"
         Top             =   1920
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select"
      Height          =   735
      Left            =   120
      TabIndex        =   17
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
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   3615
      Begin VB.TextBox txtSomething 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtMove 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtNameData 
         Height          =   885
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cboMoveType 
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
         TabIndex        =   3
         Top             =   1055
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "????"
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
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Move"
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
         TabIndex        =   20
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
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
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Move Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
   End
End
Attribute VB_Name = "Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Address As Long
Private Loading As Boolean
Private NameAddress As Long

Private Sub cboMoveType_Click()

  If Loading = False Then
   txtNameData.Visible = False
   RomDump(Address + 3) = cboMoveType.ItemData(cboMoveType.ListIndex)
  End If
End Sub

Private Sub cboName_Click()

  Dim Index As Long
  Dim Leng As Byte
  
  
  txtNameData.Visible = False
  EnableTextboxes True, Me
   
  Loading = True

  Address = GetAddress(cboName.ListIndex)
  NameAddress = GetNameAddress(cboName.ListIndex)
  
  Leng = RomDump(NameAddress - 1)
  txtName.MaxLength = Leng
  txtNameData.Text = ""
  
  For Index = 0 To Leng - 1
   If Index <> Leng - 1 Then
   txtNameData.Text = txtNameData.Text & RomDump(NameAddress + Index) & vbNewLine
   Else
   txtNameData.Text = txtNameData.Text & RomDump(NameAddress + Index)
   End If
  Next
   
  
  ' This bloc loads in the raw number data used for these names
  Dim NameArray() As String
  NameArray = Split(txtNameData.Text, vbNewLine)
   txtName.Text = ""
  For Index = 0 To UBound(NameArray())
   txtName.Text = txtName.Text & Chr(CLng(NameArray(Index)))
  Next Index
  
  
  Index = RomDump(Address + 2)
  
  chk25Fire.Value = vbUnchecked
  chk25Ice.Value = vbUnchecked
  chk25Lightning.Value = vbUnchecked
  chk25Wind.Value = vbUnchecked
  chk50Fire.Value = vbUnchecked
  chk50Ice.Value = vbUnchecked
  chk50Lightning.Value = vbUnchecked
  chk50Wind.Value = vbUnchecked
  
  If Index >= &H80 Then
   chk50Fire.Value = vbChecked
   Index = Index - &H80
  End If
  If Index >= &H40 Then
   chk25Fire.Value = vbChecked
   Index = Index - &H40
  End If
  If Index >= &H20 Then
   chk50Ice.Value = vbChecked
   Index = Index - &H20
  End If
  If Index >= &H10 Then
   chk25Ice.Value = vbChecked
   Index = Index - &H10
  End If
  If Index >= &H8 Then
   chk50Lightning.Value = vbChecked
   Index = Index - &H8
  End If
  If Index >= &H4 Then
   chk25Lightning.Value = vbChecked
   Index = Index - &H4
  End If
  If Index >= &H2 Then
   chk50Wind.Value = vbChecked
   Index = Index - &H2
  End If
  If Index >= &H1 Then
   chk25Wind.Value = vbChecked
   Index = Index - &H1
  End If
  
  SetCombo cboMoveType, CLng(RomDump(Address + 3))
  
  txtMove.Text = RomDump(Address)
  
  txtOtherResistance.Text = RomDump(Address + 1)
  txtSomething.Text = RomDump(Address + 4)
  
  '' txtName.Text = RomDump(Address + 1) & " " & RomDump(Address + 2) & " " & RomDump(Address + 3) & " " & RomDump(Address + 4)
  
  
  '' Critical chances and such crap
  
  Dim AtkChances As Byte
  
  AtkChances = RomDump(Address + 4)
  
  If CLng(AtkChances - 192) >= 0 Then
   optCounter(3).Value = True
  ElseIf CLng(AtkChances - 128) >= 0 Then
   optCounter(2).Value = True
  ElseIf CLng(AtkChances - 64) >= 0 Then
   optCounter(1).Value = True
  Else
   optCounter(0).Value = True
  End If
  
  If AtkChances - 128 >= 0 Then: AtkChances = AtkChances - 128
  If AtkChances - 64 >= 0 Then: AtkChances = AtkChances - 64
  
  If CLng(AtkChances - 48) >= 0 Then
   optDouble(3).Value = True
  ElseIf CLng(AtkChances - 32) >= 0 Then
   optDouble(2).Value = True
  ElseIf CLng(AtkChances - 16) >= 0 Then
   optDouble(1).Value = True
  Else
   optDouble(0).Value = True
  End If
  
  If AtkChances - 32 >= 0 Then: AtkChances = AtkChances - 32
  If AtkChances - 16 >= 0 Then: AtkChances = AtkChances - 16
  
  
  Dim Match As Boolean
  Dim Consider As Byte
  
  Match = False
  Consider = 15
  
  Do While Match = False
    If CLng(AtkChances - CLng(Consider)) >= 0 Then
     optCritical(Consider).Value = True
     Match = True
    End If
    
    If Consider >= 1 Then: Consider = Consider - 1
  Loop
  
  
  For AtkChances = 0 To 15
   optCritical(CInt(AtkChances)).Enabled = True
  Next AtkChances
  
  optCounter(0).Enabled = True
  optCounter(1).Enabled = True
  optCounter(2).Enabled = True
  optCounter(3).Enabled = True
  
  optDouble(0).Enabled = True
  optDouble(1).Enabled = True
  optDouble(2).Enabled = True
  optDouble(3).Enabled = True
    
  Loading = False
End Sub

Private Sub cmdRename_Click()
 Dim Rename As New NameItemClass
 Dim NameLocation As Integer

 Rename.Show
 
 NameLocation = 128 + CInt(cboName.ListIndex)
   
 Rename.cboName.ListIndex = NameLocation
End Sub

Private Sub Form_Load()

 Dim Index As Long
  
 EnableTextboxes False, Me
  
 For Index = 0 To UBound(ClassName()) '  31
   
   cboName.AddItem ClassName(Index)
   cboName.ItemData(Index) = Index
  
 Next
 
 For Index = 0 To UBound(MoveTypeName())
   cboMoveType.AddItem MoveTypeName(Index)
   cboMoveType.ItemData(Index) = MoveTypeCode(Index)
 Next
 
End Sub

Private Function GetAddress(CharacterNum As Byte) As Long

 Dim Index As Long
 Dim Count As Byte
 
'  Index = &H1EE2F0
 GetAddress = pClassData + CLng(CharacterNum) * CLng(5)

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
' 'Index = Index + 32 * 6 + 1 + CharacterNum * 5
'
' GetAddress = Index + 32 * 6 + 1 + CharacterNum * 5
 
End Function


Private Sub optCounter_Click(Index As Integer)
 If Loading = False Then: Call AssignDexterity
End Sub

Private Sub optCritical_Click(Index As Integer)
  If Loading = False Then: Call AssignDexterity
End Sub

Private Sub optDouble_Click(Index As Integer)
 If Loading = False Then: Call AssignDexterity
End Sub

Private Sub txtMove_Change()

 Dim Result As Byte
 
 If ValidateInput(txtMove, False, True) = True Then
  Result = txtMove.Text
  RomDump(Address) = Result
 End If
 
End Sub

Private Sub txtMove_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtMove_LostFocus()

 Dim Result As Byte
 
 If ValidateInput(txtMove, False) = True Then
  Result = txtMove.Text
  RomDump(Address) = Result
 End If
 
End Sub

Private Sub txtName_Change()
'' Dim Index As Long
'' Dim Result As String
''
'' If Loading = False Then
''
'' For Index = 1 To Len(txtName.Text)
''
''  Result = Result & ChCode(Mid(txtName, Index, 1)) & vbNewLine
''
'' Next Index
''
'' Index = 0
''
'' Do While Len(txtName.Text) + Index < txtName.MaxLength
''
''  Result = Result & "32" & vbNewLine
''
''  Index = Index + 1
'' Loop
''
'' txtNameData.Text = Left(Result, Len(Result) - 2)
''
'' End If
End Sub

Private Sub txtName_Click()
'' txtNameData.Visible = True
End Sub


Private Sub txtNameData_Change()
''  On Error GoTo Skipo
''
''  Dim Index As Long
''
''  Dim NameArray() As String
''
''  Dim Result As String
''
''  NameArray = Split(txtNameData.Text, vbNewLine)
''
''  For Index = 0 To UBound(NameArray())
''   Result = Result & Chr(CLng(CByte(NameArray(Index))))
''  Next Index
''
''  For Index = 0 To UBound(NameArray())
''
''   RomDump(NameAddress + Index) = CByte(NameArray(Index))
''
''  Next Index
''
''
''  Exit Sub
''
''Skipo:
End Sub

Private Sub txtOtherResistance_Change()
 Dim Result As Byte
 
 If ValidateInput(txtOtherResistance, False, True) = True Then
  Result = txtOtherResistance.Text
  RomDump(Address + 1) = Result
 End If
End Sub

Private Sub txtOtherResistance_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtOtherResistance_LostFocus()
 
 Dim Result As Byte
 
 If ValidateInput(txtOtherResistance, False) = True Then
  Result = txtOtherResistance.Text
  RomDump(Address + 1) = Result
 End If
 
End Sub

Private Sub txtSomething_Change()
 Dim Result As Byte
 
 If ValidateInput(txtSomething, False, True) = True Then
  Result = txtSomething.Text
  RomDump(Address + 4) = Result
 End If
End Sub

Private Sub txtSomething_Click()
 txtNameData.Visible = False
End Sub

Private Sub txtSomething_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtSomething, False) = True Then
  Result = txtSomething.Text
  RomDump(Address + 4) = Result
 End If
End Sub



Private Sub chk25Ice_Click()
 
 Dim Sum As Byte
  
 If Loading = True Then
  Exit Sub
 End If
 
 txtNameData.Visible = False
 
 If chk25Fire.Value = vbChecked Then
  Sum = Sum + chk25Fire.Tag
 End If
 If chk50Fire.Value = vbChecked Then
  Sum = Sum + chk50Fire.Tag
 End If
 If chk25Ice.Value = vbChecked Then
  Sum = Sum + chk25Ice.Tag
 End If
 If chk50Ice.Value = vbChecked Then
  Sum = Sum + chk50Ice.Tag
 End If
 If chk25Wind.Value = vbChecked Then
  Sum = Sum + chk25Wind.Tag
 End If
 If chk50Wind.Value = vbChecked Then
  Sum = Sum + chk50Wind.Tag
 End If
 If chk25Lightning.Value = vbChecked Then
  Sum = Sum + chk25Lightning.Tag
 End If
 If chk50Lightning.Value = vbChecked Then
  Sum = Sum + chk50Lightning.Tag
 End If
 
 RomDump(Address + 2) = Sum
End Sub

Private Sub chk25Fire_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk25Lightning_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk25Wind_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk50Fire_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk50Ice_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk50Lightning_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub chk50Wind_Click()
 Call chk25Ice_Click
 txtNameData.Visible = False
End Sub

Private Sub txtNameData_KeyDown(KeyCode As Integer, Shift As Integer)
 
 DoEvents
 
 If KeyCode = 13 Then
  txtNameData.Text = Left(txtNameData.Text, txtNameData.SelStart - 2) & Right(txtNameData.Text, Len(txtNameData.Text) - txtNameData.SelStart)
 End If
End Sub

Private Function GetNameAddress(ByVal iName As Long) As Long

 Dim Index As Long
 Dim Count As Long
 
 Count = CLASSNAMES_ORIGINAL_OFFSET

 Do While Index < iName
  Count = Count + RomDump(Count) + 1
  
  Index = Index + 1
 Loop
 
 GetNameAddress = Count + 1
 
End Function

Private Function AssignDexterity()

 Dim Index As Integer
 Dim Value As Byte
   
  Value = 0
  
  Select Case True
  Case optCounter(3).Value
   Value = Value + 192
  Case optCounter(2).Value
   Value = Value + 128
  Case optCounter(1).Value
   Value = Value + 64
  Case optCounter(0).Value
   Value = Value + 0
  End Select
  
  Select Case True
  Case optDouble(3).Value
   Value = Value + 48
  Case optDouble(2).Value
   Value = Value + 32
  Case optDouble(1).Value
   Value = Value + 16
  Case optDouble(0).Value
   Value = Value + 0
  End Select
  
  For Index = 0 To 15
   If optCritical(Index).Value = True Then: Value = Value + CByte(Index)
  Next Index


  RomDump(Address + 4) = Value

End Function
