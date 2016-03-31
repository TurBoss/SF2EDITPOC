VERSION 5.00
Begin VB.MDIForm Main 
   BackColor       =   &H8000000C&
   Caption         =   "SF2 Editor"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8385
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRomEditors 
      Caption         =   "Rom Editors"
      Enabled         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Begin VB.Menu mnuCharacterData 
            Caption         =   "&Character Start Data"
         End
         Begin VB.Menu mnuCharacterStats 
            Caption         =   "Character S&tats"
         End
         Begin VB.Menu mnuExpandStat 
            Caption         =   "Expand Stat Table"
         End
         Begin VB.Menu Sep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClass 
            Caption         =   "&Class"
         End
         Begin VB.Menu mnuLevels 
            Caption         =   "&Levels"
         End
         Begin VB.Menu mnuPromotions 
            Caption         =   "&Promotions"
         End
         Begin VB.Menu Sep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMonsters 
            Caption         =   "&Monsters"
         End
         Begin VB.Menu Sep5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuItems 
            Caption         =   "&Items"
         End
         Begin VB.Menu mnuShops 
            Caption         =   "&Shops"
         End
         Begin VB.Menu Sep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSpells 
            Caption         =   "Sp&ells"
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "M&isc"
         End
      End
      Begin VB.Menu mnuEditNames 
         Caption         =   "Edit Names"
         Enabled         =   0   'False
         Begin VB.Menu mnuItemClass 
            Caption         =   "Items && Class"
         End
         Begin VB.Menu mnuSpellHeroMonster 
            Caption         =   "Spells, Heroes && Monsters"
         End
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert"
         Enabled         =   0   'False
         Begin VB.Menu mnuSMD2BIN 
            Caption         =   "SMD -> BIN"
         End
      End
   End
   Begin VB.Menu mnuDisasmEditors 
      Caption         =   "Disasm Editors"
      Enabled         =   0   'False
      Begin VB.Menu mnuNames 
         Caption         =   "Names"
         Begin VB.Menu mnuDisasmItemNames 
            Caption         =   "Item Names"
         End
      End
   End
   Begin VB.Menu mnuMisc 
      Caption         =   "Misc"
      Enabled         =   0   'False
      Begin VB.Menu mnuStatCalculator 
         Caption         =   "Stat Calculator"
      End
      Begin VB.Menu mnuDamageCalculator 
         Caption         =   "Damage Calculator"
      End
      Begin VB.Menu mnuFixCharPointTable 
         Caption         =   "Fix Pointers"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuUsers 
         Caption         =   "Users"
      End
      Begin VB.Menu mnuDevs 
         Caption         =   "Developers"
      End
      Begin VB.Menu Sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
 Load_Codes
 Load_Shops
 Load_Path
 
 Randomize
 
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 If RomPath <> "" And disasmMode = False Then
  If MsgBox("You've loaded a ROM file, would you like to save any changes you may have made to it?", vbYesNo) = vbYes Then
    Call mnuSave_Click
  End If
 Else
  If DisasmConfFilePath <> "" And disasmMode = True Then
    If MsgBox("You've loaded a Disassembly, would you like to save any changes you may have made to its files?", vbYesNo) = vbYes Then
      Call mnuSave_Click
    End If
  End If
 End If
End Sub

Private Sub mnuAbout_Click()
 Dim blah As New Reader
 
 blah.Caption = "About"
 
 blah.Show
End Sub

Private Sub mnuCharacterData_Click()
 Dim blah As New CharacterData
 
 blah.Show
End Sub

Private Sub mnuCharacterStats_Click()
  
 If CharacterStats.Tag = 1 Then
 
  CharacterStats.Tag = 0
 
 Else
 
  Dim blah As New CharacterStats

  blah.Show
  
 End If
  
End Sub

Private Sub mnuDamageCalculator_Click()
 Dim blah As New DamageCalculator
 
 blah.Show
End Sub


Private Sub mnuClass_Click()
 Dim blah As New Class
 
 blah.Show
End Sub

Private Sub mnuDevs_Click()
 Dim blah As New Reader
 
 blah.Caption = "Developers"
 
 blah.Show
End Sub

Private Sub mnuStatCalculator_Click()
 Dim blah As New StatCalculator
 
 blah.Show
End Sub


Private Sub mnuExit_Click()
  
  'Unload All
  
  End
End Sub

Private Sub mnuExpandStat_Click()
 
 InsertStats.Show
 
End Sub

Private Sub mnuFixCharPointTable_Click()

 Dim Index As Long
 Dim Counter As Integer

 Dim MetaCounter As Integer
 
 Counter = 0
 MetaCounter = 1
 
 Index = ALLYSTATSDATASTART_ORIGINAL_OFFSET

 Do While Counter < 29


  Do While MetaCounter <> 3 And MetaCounter <> 6 _
     And MetaCounter <> 9 And MetaCounter <> 12 _
     And MetaCounter <> 15 And MetaCounter <> 17 _
     And MetaCounter <> 19 And MetaCounter <> 21 _
     And MetaCounter <> 23 And MetaCounter <> 25 _
     And MetaCounter <> 27 And MetaCounter <> 28 _
     And MetaCounter <> 31 And MetaCounter <> 34 _
     And MetaCounter <> 37 And MetaCounter <> 40 _
     And MetaCounter <> 43 And MetaCounter <> 46 And MetaCounter < 49 _
    
      
   Do While RomDump(Index) < 254
    Index = Index + 1
   Loop
   
   MetaCounter = MetaCounter + 1

  Index = Index + 1
  
  Loop

  Counter = Counter + 1

  RomDump(ALLYSTATS_ORIGINAL_OFFSET + 1 + 4 * Counter) = CByte(Fix(Index / 65536#))
  RomDump(ALLYSTATS_ORIGINAL_OFFSET + 1 + 4 * Counter + 1) = Fix((Index - Fix(Index / 65536#) * 65536#) / 256#)
  RomDump(ALLYSTATS_ORIGINAL_OFFSET + 1 + 4 * Counter + 2) = CByte(Index - Fix(Fix(Index / 256#) * 256#))

  Do While RomDump(Index) < 254
   Index = Index + 1
  Loop
   
  MetaCounter = MetaCounter + 1
  
  Index = Index + 1
 Loop


 '' Update joining guy data pointers
 Index = Index + 1
 
 RomDump(ALLYSTARTDATAPOINTER_ORIGINAL_OFFSET + 1) = CByte(Fix(Index / 65536#))
 RomDump(ALLYSTARTDATAPOINTER_ORIGINAL_OFFSET + 1 + 1) = Fix((Index - Fix(Index / 65536#) * 65536#) / 256#)
 RomDump(ALLYSTARTDATAPOINTER_ORIGINAL_OFFSET + 1 + 2) = CByte(Index - Fix(Fix(Index / 256#) * 256#))

 'Update Class Pointer
 
 Index = Index + 192

 RomDump(CLASSDATAPOINTER_ORIGINAL_OFFSET + 1) = CByte(Fix(Index / 65536#))
 RomDump(CLASSDATAPOINTER_ORIGINAL_OFFSET + 1 + 1) = Fix((Index - Fix(Index / 65536#) * 65536#) / 256#)
 RomDump(CLASSDATAPOINTER_ORIGINAL_OFFSET + 1 + 2) = CByte(Index - Fix(Fix(Index / 256#) * 256#))
 
 'Update 01010101 Pointer
 
 Index = Index + 155
 
 RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1) = CByte(Fix(Index / 65536#))
 RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 1) = Fix((Index - Fix(Index / 65536#) * 65536#) / 256#)
 RomDump(JEWELENDSCREENLAYOUTPOINTER_ORIGINAL_OFFSET + 1 + 2) = CByte(Index - Fix(Fix(Index / 256#) * 256#))

  
  MsgBox "Pointer table adjusted.", vbOKOnly
End Sub

Private Sub mnuItemClass_Click()
 Dim blah As New NameItemClass
 
 blah.Show
End Sub

Private Sub mnuItems_Click()
 Dim blah As New Item
 
 blah.Show
End Sub

Private Sub mnuLevels_Click()
 Dim blah As New Levels
 
 blah.Show
End Sub

Private Sub mnuMonsters_Click()
 Dim blah As New Monsters
 
 blah.Show
End Sub

Private Sub mnuOpen_Click()
  SelectFile.Show vbModal
End Sub

Private Sub mnuPromotions_Click()
 Dim blah As New Promotions
 
 blah.Show
End Sub

Private Sub mnuSave_Click()

 If RomPath <> "" And disasmMode = False Then
    Call WriteRom
 Else
  If DisasmConfFilePath <> "" And disasmMode = True Then
    Call WriteDisasmFiles
  End If
 End If

End Sub


Private Sub mnuShops_Click()
 Dim blah As New Shops
 
 blah.Show
End Sub

Private Sub mnuSpellLists_Click()
 Dim blah As New SpellLists
 
 blah.Show
End Sub

Private Sub mnuSMD2BIN_Click()
  Convert.Show vbModal
End Sub

Private Sub mnuSpecial_Click()
 Dim blah As New Special
 
 blah.Show
End Sub

Private Sub mnuSpellHeroMonster_Click()
 Dim blah As New NameSpellHeroMonster
 
 blah.Show
End Sub

Private Sub mnuSpells_Click()
 Dim blah As New Spells
  
 blah.Show
 
 'SpellLook.Show
End Sub

Private Sub mnuUsers_Click()
 
 Dim blah As New Reader
 
 blah.Caption = "Users"
 
 blah.Show
 
End Sub
