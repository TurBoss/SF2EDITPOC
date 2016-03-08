VERSION 5.00
Begin VB.Form Convert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMD -> BIN Converter"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5070
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdSave_Click()
       
  Dim Ein As Long
  Dim Oan As Long
  
  Dim Index As Long
  Dim MetaIndex As Long
  
  Dim Blocks As Long
  
  Ein = 0
  Oan = 1

  'Blocks = Fix((UBound(RomDump()) + 1 - 512) / 16384)
   Blocks = Fix((&H1FFFFF + 1 - 512) / 16384)
               
  For MetaIndex = 0 To Blocks
  
  For Index = 0 To 16383
    If Index < 8192 Then
      RomConverted(Oan) = RomUnConverted(Index + MetaIndex * 16384 + 512)
      Oan = Oan + 2
    Else
      RomConverted(Ein) = RomUnConverted(Index + MetaIndex * 16384 + 512)
      Ein = Ein + 2
    End If
  Next Index
  
  Next MetaIndex

  
  For Index = 0 To UBound(RomDump())
   RomDump(Index) = RomConverted(Index)
  Next Index
       
       
  Dim Freedfile As Long

 Freedfile = FreeFile()

 
 Rom.RomPath = Text1.Text

 Open RomPath For Binary As #Freedfile
  
  Put #1, , RomDump

 Close #Freedfile

 Call InitializeAddresses
 GuyNumber = CountGuys()
  
 MsgBox "Saved~!", vbOKOnly
 
 
 
 Dim Count As Long
 Dim SubIndex As Long
 
 Index = pItemNames '&H1796E
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

 Index = pSpellNames  '63940

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
 
End Sub

Private Sub Form_Load()
  
  Dim Index As Long
  
  Dim Freedfile As Long
  
 Freedfile = FreeFile()
 
 Open Rom.RomPath For Binary As #Freedfile
  Get #Freedfile, , RomUnConverted
 Close #Freedfile
 
   
  If Len(Rom.RomPath) > 4 Then
    Text1.Text = Left(Rom.RomPath, Len(Rom.RomPath) - 4) & ".bin"
  Else
    Text1.Text = Rom.RomPath
  End If

End Sub
