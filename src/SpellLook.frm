VERSION 5.00
Begin VB.Form SpellLook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   10620
   Begin VB.ComboBox cbospell 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   10215
   End
End
Attribute VB_Name = "SpellLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ParseLength As Long = 64 * 100


Private Sub cbospell_Click()
 Dim Index As Long
 Dim SubIndex As Long
 Dim Stringy(ParseLength) As String
 Dim DisplayString As String
 
 Dim Address As Long
 ' Wiz note : 0x177BF is inside Spell Defs (starts at 0x176A6) ... I don't get it.
 Address = &H177BF - 1 + 32 - 32 * 206 + 32 * cboSpell.ItemData(cboSpell.ListIndex)

 For Index = Address To Address + 31
 
  DisplayString = DisplayString & Align(CByte(RomDump(Index))) & RomDump(Index) & " "
  
 Next
 
 Text1.Text = DisplayString
End Sub

Private Sub Form_Load()
 Dim Index As Long
 Dim SubIndex As Long
 Dim Stringy(ParseLength) As String
 Dim DisplayString As String
 
 Dim Address As Long
 
 For Index = 0 To UBound(SpellName())
  cboSpell.AddItem SpellName(Index)
  cboSpell.ItemData(Index) = SpellCode(Index)
 Next
 
 Address = &H177BF - 1 + 64 * 10
 
 
 For Index = 0 To ParseLength
 
  Stringy(Index) = RomDump(Address - Index)
   
 Next Index
 
 For Index = ParseLength To 0 Step -1
 
  If Index / 8# = Fix(Index / 8#) Then
   DisplayString = DisplayString & vbNewLine & "--------------" & vbNewLine
  End If
    
  DisplayString = DisplayString & Align(CByte(Stringy(Index))) & Stringy(Index) & " "
  
''  If Index / 16# = Fix(Index / 16#) And Index / 32# <> Fix(Index / 32#) Then
''   DisplayString = DisplayString & vbNewLine
''  End If
 Next
 
 Text1.Text = DisplayString
 
End Sub

Private Function Align(poo As Byte) As String

 If poo < 10 Then
  Align = "  "
 ElseIf poo < 100 Then
  Align = " "
 Else
  Align = ""
 End If

End Function
