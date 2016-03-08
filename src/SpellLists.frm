VERSION 5.00
Begin VB.Form SpellLists 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Lists"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   8505
   Begin VB.ListBox List1 
      Height          =   7275
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "SpellLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

 Dim index As Long
 Dim subindex As Long
 
 For index = 0 To 2000
 
  subindex = 0
  
  List1.AddItem index & " - " & RomDump(&H1EE300 + index)
  
  index = index + 1
  
  Do While SpellCode(subindex) <> RomDump(&H1EE300 + index) And subindex < UBound(SpellCode())
  
   subindex = subindex + 1
   
  Loop
 
  If subindex <= UBound(SpellName()) Then
   List1.AddItem index & " - " & SpellName(subindex)
  Else
   List1.AddItem index & " - " & "  "
  End If
 
 Next index
 

End Sub
