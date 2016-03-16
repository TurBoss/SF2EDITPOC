VERSION 5.00
Begin VB.Form InsertStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stat Block Insertion"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   5445
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Slot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Slot"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   2
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ListBox lstBlocks 
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
         TabIndex        =   1
         Top             =   330
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Close any open character stats screen before adding or removing stat blocks."
         Height          =   1335
         Left            =   3120
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Stat Blocks:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "InsertStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 
 Dim Index As Long
 Dim SubIndex As Long
 
' For Index = 1 To GuyBlocks(0)
'   lstBlocks.AddItem CharacterName(0) & " " & Index
' Next Index
'
' For Index = 1 To 29
'   For SubIndex = GuyBlocks(Index - 1) To GuyBlocks(Index) - 1
'    lstBlocks.AddItem CharacterName(Index) & " " & SubIndex - GuyBlocks(Index - 1) + 1
'   Next SubIndex
' Next

  
 For Index = 0 To UBound(GuyBlocks())   '29
   lstBlocks.AddItem PersonName(Index)  '  CharacterName(Index)
 Next Index
End Sub

Private Sub lstBlocks_Click()

 Dim HowMany As Integer
 
 
 If lstBlocks.ListIndex = 0 Then
  HowMany = GuyBlocks(lstBlocks.ListIndex)
 Else
  HowMany = GuyBlocks(lstBlocks.ListIndex) - GuyBlocks(lstBlocks.ListIndex - 1)
 End If
 
 Label1.Caption = "Stat Blocks: " & HowMany

End Sub

Private Sub cmdAdd_Click()

 Dim SpriteAddress As Long
 Dim Addblock As Long

 If lstBlocks.ListIndex < 0 Then
  Exit Sub
 End If
 
 If lstBlocks.ListIndex > 0 Then
  Addblock = GuyBlocks(lstBlocks.ListIndex) - GuyBlocks(lstBlocks.ListIndex - 1) + 1
 Else
  Addblock = GuyBlocks(lstBlocks.ListIndex) + 1
 End If
 
 ' Adding in a battle sprite here
 If Addblock < 4 Then

  ' SpriteAddress = &H1F806 + lstBlocks.ListIndex * 9 + (Addblock - 1) * 3
  SpriteAddress = pBattleSprites + lstBlocks.ListIndex * 9 + (Addblock - 1) * 3
  
  RomDump(SpriteAddress) = 0
  RomDump(SpriteAddress + 1) = 0
  
 
 
  Dim ClassAddress As Long
  
  ClassAddress = pClassData   'CLng(CLng(RomDump(&H1EE00D)) * 65536 + CLng(RomDump(&H1EE00D + 1)) * 256 + CLng(RomDump(&H1EE00D + 2)))
 
  
  Dim JoinDataAddress As Long
 
  JoinDataAddress = pJoinData 'CLng(CLng(RomDump(&H1EE009)) * 65536 + CLng(RomDump(&H1EE009 + 1)) * 256 + CLng(RomDump(&H1EE009 + 2)))
 
  
  
  Dim Index As Long
  Dim SubIndex As Long
  Dim StopIndex As Long ' Be sure to write in the dummy block here
 
 
  ' Counting blocks until the new one
  
  Index = GuyPointers(lstBlocks.ListIndex)
  SubIndex = 1
  
  
  Do While SubIndex <> Addblock
   
   If RomDump(Index) = 255 Or RomDump(Index) = 254 Then
    SubIndex = SubIndex + 1
   End If
   
   Index = Index + 1
  Loop
    
  
  
  StopIndex = Index
 
  SubIndex = ClassAddress + 160
 
 
  For Index = SubIndex To StopIndex Step -1
   RomDump(Index + 17) = RomDump(Index)
  Next
 
 
  For Index = StopIndex To StopIndex + 15
   RomDump(Index) = 0
  Next Index
  
  RomDump(StopIndex + 16) = 255
  
  
  'save the new pointer addresses
  
  
  Dim GuyAdded As Long
  Dim Result As Byte
  Dim WorkingNumber As Long
  
  GuyAdded = lstBlocks.ListIndex
  
  If GuyAdded < UBound(GuyBlocks()) Then    '29 Then
   For Index = GuyAdded + 1 To UBound(GuyBlocks())   '  29
    GuyPointers(Index) = GuyPointers(Index) + 17
    
    Result = CByte(Fix(GuyPointers(Index) / 65536#))
    RomDump(ALLYSTATS_ORIGINAL_OFFSET + 1 + Index * 4) = Result

    WorkingNumber = CLng(GuyPointers(Index) - (65536 * CLng(Result)))

    Result = CByte(Fix(WorkingNumber / 256#))
    RomDump(ALLYSTATS_ORIGINAL_OFFSET + 1 + Index * 4 + 1) = Result

    WorkingNumber = CLng(WorkingNumber - (256 * CLng(Result)))
    RomDump(ALLYSTATS_ORIGINAL_OFFSET + 1 + Index * 4 + 2) = CByte(WorkingNumber)
       
   Next
  End If
  
  
  
  JoinDataAddress = JoinDataAddress + 17
  pJoinData = JoinDataAddress
  
    Result = CByte(Fix(JoinDataAddress / 65536#))
    RomDump(ALLYSTARTDATAPOINTER_ORIGINAL_OFFSET + 1) = Result

    WorkingNumber = CLng(JoinDataAddress - (65536 * CLng(Result)))

    Result = CByte(Fix(WorkingNumber / 256#))
    RomDump(ALLYSTARTDATAPOINTER_ORIGINAL_OFFSET + 1 + 1) = Result

    WorkingNumber = CLng(WorkingNumber - (256 * CLng(Result)))
    RomDump(ALLYSTARTDATAPOINTER_ORIGINAL_OFFSET + 1 + 2) = CByte(WorkingNumber)
  
    
  
  ClassAddress = ClassAddress + 17
  pClassData = ClassAddress
  
    Result = CByte(Fix(ClassAddress / 65536#))
    RomDump(CLASSDATAPOINTER_ORIGINAL_OFFSET + 1) = Result

    WorkingNumber = CLng(ClassAddress - (65536 * CLng(Result)))

    Result = CByte(Fix(WorkingNumber / 256#))
    RomDump(CLASSDATAPOINTER_ORIGINAL_OFFSET + 1 + 1) = Result

    WorkingNumber = CLng(WorkingNumber - (256 * CLng(Result)))
    RomDump(CLASSDATAPOINTER_ORIGINAL_OFFSET + 1 + 2) = CByte(WorkingNumber)
  
     
  
 
 
 Dim Roundover As Byte
 
 'Update 01010101 Pointer
 If RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) < 239 Then
   RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) = RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) + 17
 Else
   Roundover = CByte(17 - (255 - RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2)))
 
   RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) = Roundover
      
   RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 1) = RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 1) + 1
 End If
  
  ReloadPStats
  GuyNumber = CountGuys()
  
  Call lstBlocks_Click
  
 End If

 

End Sub


Private Sub cmdRemove_Click()

 Dim SpriteAddress As Long
 Dim Removeblock As Long
 Dim Blocklength As Long


 If lstBlocks.ListIndex > 0 Then
  Removeblock = GuyBlocks(lstBlocks.ListIndex) - GuyBlocks(lstBlocks.ListIndex - 1) + 1
 Else
  Removeblock = GuyBlocks(lstBlocks.ListIndex) + 1
 End If
 
 ' Removing the battle sprite here
 If Removeblock > 2 Then

  ' SpriteAddress = &H1F806 + lstBlocks.ListIndex * 9 + (Removeblock - 2) * 3
  SpriteAddress = pBattleSprites + lstBlocks.ListIndex * 9 + (Removeblock - 2) * 3
  
  
  RomDump(SpriteAddress) = 255
  RomDump(SpriteAddress + 1) = 255
  RomDump(SpriteAddress + 2) = 0
  
  
  'delete everything below this if it doesn't work
  
  Dim ClassAddress As Long
  
  ClassAddress = pClassData 'CLng(CLng(RomDump(&H1EE00D)) * 65536 + CLng(RomDump(&H1EE00D + 1)) * 256 + CLng(RomDump(&H1EE00D + 2)))
 
  
  Dim JoinDataAddress As Long
 
  JoinDataAddress = pJoinData  ' CLng(CLng(RomDump(&H1EE009)) * 65536 + CLng(RomDump(&H1EE009 + 1)) * 256 + CLng(RomDump(&H1EE009 + 2)))
 
  
  
  Dim Index As Long
  Dim SubIndex As Long
  Dim StopIndex As Long ' Be sure to write in the dummy block here
 
 
  ' Counting blocks until the new one
  
  Index = GuyPointers(lstBlocks.ListIndex)
  SubIndex = 1
  
  
  Do While SubIndex <> Removeblock
   
   If RomDump(Index) = 255 Or RomDump(Index) = 254 Then
    SubIndex = SubIndex + 1
    
    ' need to chop off the spell lists, too.
    If SubIndex + 1 = Removeblock Then
     Blocklength = Index
    End If
   End If
   
   Index = Index + 1
  Loop
  
  Blocklength = Index - Blocklength - 1
  
  StopIndex = Index
 
  SubIndex = ClassAddress + 160
 
 
  ' Do this from the front to the back silly man
  For Index = StopIndex To SubIndex Step 1
   RomDump(Index - Blocklength) = RomDump(Index)
  Next
 
''  Don't need an empty stat block cause it's being overwritten
''
''  For Index = StopIndex To StopIndex + 15
''   RomDump(Index) = 0
''  Next Index
''
''  RomDump(StopIndex + 16) = 255
  
  
  'save the new pointer addresses
  
  
  Dim GuyAdded As Long
  Dim Result As Byte
  Dim WorkingNumber As Long
  
  GuyAdded = lstBlocks.ListIndex
  
  ' Because it only affects everyone above him on the list
  If GuyAdded < UBound(GuyBlocks()) Then '29 Then
   For Index = GuyAdded + 1 To 29
    GuyPointers(Index) = GuyPointers(Index) - Blocklength
    
    Result = CByte(Fix(GuyPointers(Index) / 65536#))
    RomDump(ALLYSTATS_ORIGINAL_OFFSET + 1 + Index * 4) = Result

    WorkingNumber = CLng(GuyPointers(Index) - (65536 * CLng(Result)))

    Result = CByte(Fix(WorkingNumber / 256#))
    RomDump(ALLYSTATS_ORIGINAL_OFFSET + 1 + Index * 4 + 1) = Result

    WorkingNumber = CLng(WorkingNumber - (256 * CLng(Result)))
    RomDump(ALLYSTATS_ORIGINAL_OFFSET + 1 + Index * 4 + 2) = CByte(WorkingNumber)
       
   Next
  End If
  
  
  
  JoinDataAddress = JoinDataAddress - Blocklength
  pJoinData = JoinDataAddress
  
    Result = CByte(Fix(JoinDataAddress / 65536#))
    RomDump(ALLYSTARTDATAPOINTER_ORIGINAL_OFFSET + 1) = Result

    WorkingNumber = CLng(JoinDataAddress - (65536 * CLng(Result)))

    Result = CByte(Fix(WorkingNumber / 256#))
    RomDump(ALLYSTARTDATAPOINTER_ORIGINAL_OFFSET + 1 + 1) = Result

    WorkingNumber = CLng(WorkingNumber - (256 * CLng(Result)))
    RomDump(ALLYSTARTDATAPOINTER_ORIGINAL_OFFSET + 1 + 2) = CByte(WorkingNumber)
  
    
  
  ClassAddress = ClassAddress - Blocklength
  pClassData = ClassAddress
  
    Result = CByte(Fix(ClassAddress / 65536#))
    RomDump(CLASSDATAPOINTER_ORIGINAL_OFFSET + 1) = Result

    WorkingNumber = CLng(ClassAddress - (65536 * CLng(Result)))

    Result = CByte(Fix(WorkingNumber / 256#))
    RomDump(CLASSDATAPOINTER_ORIGINAL_OFFSET + 1 + 1) = Result

    WorkingNumber = CLng(WorkingNumber - (256 * CLng(Result)))
    RomDump(CLASSDATAPOINTER_ORIGINAL_OFFSET + 1 + 2) = CByte(WorkingNumber)
  
     
  
 
 
 Dim Roundover As Byte
 
 If isExpanded = False Then
 'Update 01010101 Pointer
 If isExpanded = False Then
 If RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) > Blocklength - 1 Then
   RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) = RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) - Blocklength
 Else
   Roundover = CByte((255 + RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) - Blocklength))
 
   RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) = Roundover
      
   RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 1) = RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 1) - 1
 End If
 End If
  
 End If
  
  ReloadPStats
  GuyNumber = CountGuys()
    
  Call lstBlocks_Click
  
  
 End If


End Sub
