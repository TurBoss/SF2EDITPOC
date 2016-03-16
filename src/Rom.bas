Attribute VB_Name = "Rom"

Public RomDump() As Byte
Public RomPath As String

Public RomConverted(&H1FFFFF) As Byte
Public RomUnConverted(&H1FFFFF + 512) As Byte

Public DefaultPath As String

Public pItemData As Long
Public pMonsterData As Long
Public pClassData As Long
Public pBattleSprites As Long
Public pPromotions As Long
Public pSpells As Long

Public pItemNames As Long
Public pSpellNames As Long

Public pStats() As Long

Public pJoinData As Long

Public isExpanded As Boolean
Public disasmMode As Boolean

' Just input the first byte of the pointer and it'll load the rest
Public Function LoadPointer(PointerAddress As Long) As Long

Dim Result As Long

' Result = Result + CLng(RomDump(PointerAddress)) * CLng(16777216)
Result = Result + CLng(RomDump(PointerAddress + 1)) * CLng(65536)
Result = Result + CLng(RomDump(PointerAddress + 2)) * CLng(256)
Result = Result + CLng(RomDump(PointerAddress + 3))

LoadPointer = Result

' 256
' 65536
' 16777216

End Function

Public Sub SetPointer(PointerAddress As Long, NewValue As Long)

  Dim Result As Long
  Dim WorkingNumber As Long

    Result = CByte(Fix(NewValue / 65536#))
    RomDump(PointerAddress + 1) = Result

    WorkingNumber = CLng(NewValue - (65536 * CLng(Result)))

    Result = CByte(Fix(WorkingNumber / 256#))
    RomDump(PointerAddress + 2) = Result

    WorkingNumber = CLng(WorkingNumber - (256 * CLng(Result)))
    RomDump(PointerAddress + 3) = CByte(WorkingNumber)


' 256
' 65536
' 16777216

End Sub

Public Sub InitializeAddresses()

  Dim Index As Long
  Dim aLong As Long
  Dim wQuit As Boolean
  
  pMonsterData = ENEMYDATA_ORIGINAL_OFFSET
  pItemData = ITEMDEFS_ORIGINAL_OFFSET
  pClassData = CLASSDATA_ORIGINAL_OFFSET
  pBattleSprites = BATTLESPRITES_ORIGINAL_OFFSET

  pPromotions = PROMOTIONS_ORIGINAL_OFFSET
  pSpells = SPELLDEFS_ORIGINAL_OFFSET
  pItemNames = ITEMNAMES_ORIGINAL_OFFSET
  pSpellNames = SPELLNAMES_ORIGINAL_OFFSET
  
  pJoinData = ALLYSTARTDATA_ORIGINAL_OFFSET

  ' Gather up each pointer to each "dude"
  Index = 0
  
  Do
  
   ReDim Preserve pStats(Index)
        
   aLong = ALLYSTATS_ORIGINAL_OFFSET + 4 * Index
        
   pStats(Index) = LoadPointer(aLong)
   
   If LoadPointer(aLong + 4) - LoadPointer(aLong) > 150 Or LoadPointer(aLong + 4) < LoadPointer(aLong) Then
     wQuit = True
   End If
    
   If LoadPointer(aLong + 4) = 2025406 And Index = 29 Then
     wQuit = True
   
     SetPointer aLong + 4, 0
     SetPointer aLong + 8, 0
   End If
    
   Index = Index + 1
    
  Loop While wQuit = False
  

  For Index = 0 To UBound(ItemOffset)
   ItemOffset(Index) = pItemData + 16 * Index
  Next Index

  isExpanded = False

  If LoadPointer(2023444) < 1888000 Then ' jewel end screen layout pointer
   isExpanded = True
  End If

  GuyNumber = CountGuys()
  
  'Main.Caption = UBound(pStats())
End Sub

Public Sub ReloadPStats()
  
  Dim Index As Long
  Dim aLong As Long
  
  Index = 0
  
 Do
  
' For Index = 0 To UBound(pStats())
   aLong = ALLYSTATS_ORIGINAL_OFFSET + 4 * Index
        
   pStats(Index) = LoadPointer(aLong)
   Index = Index + 1
   
   If LoadPointer(aLong + 4) - LoadPointer(aLong) > 150 Or LoadPointer(aLong + 4) < LoadPointer(aLong) Then
     wQuit = True
   End If
' Next Index
    
  Loop While wQuit = False

End Sub



