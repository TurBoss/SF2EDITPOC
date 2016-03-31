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

Public Sub LoadRom(FileAbsolutePath As String)

On Error GoTo OnError

     Dim Freedfile As Long
     Freedfile = FreeFile()
 
     Open FileAbsolutePath For Binary As #Freedfile
      If UCase(Right(FileAbsolutePath, 3)) <> "SMD" Then  'And Right(File1.FileName, 3) <> "SMD" Then
       ReDim RomDump(LOF(Freedfile) - 1)
      Else
       ReDim RomDump(&H1FFFFF)
      End If
      Get #Freedfile, , RomDump
     Close #Freedfile
    
     If UCase(Right(FileAbsolutePath, 3)) <> "SMD" Then
      Call InitializeAddresses
     End If
 
    ' Do stuff we couldn't before load
    CalculateStoreSpots
    
    RomPath = FileAbsolutePath
    
    Main.mnuDisasmEditors.Enabled = False
    Main.mnuRomEditors.Enabled = True
    Main.mnuEdit.Enabled = True
    Main.mnuConvert.Enabled = True
    Main.mnuMisc.Enabled = True
    Main.mnuEditNames.Enabled = True
    
    Dim Index As Long
    Dim Count As Long
    Dim SubIndex As Long
    
    Index = pItemNames    '&H1796E
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
    Index = pSpellNames '63940
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
     
    ''' If UCase(Right(File1.FileName, 3)) <> "SMD" Then 'And Right(File1.FileName, 3) <> "SMD" Then
    '''  GuyNumber = CountGuys()
    ''' End If
    Call LoadRomNames

Exit Sub

OnError:
    Close Freedfile
    MsgBox "Error while loading file.", vbOKOnly
    
End Sub


Public Sub InitializeAddresses()

  Dim Index As Long
  Dim aLong As Long
  Dim wQuit As Boolean
  
  pMonsterData = ENEMYDATA_ORIGINAL_OFFSET
  pItemData = ITEMDEFS_ORIGINAL_OFFSET
  pClassData = CLASSDATA_ORIGINAL_OFFSET
  pBattleSprites = ALLYBATTLESPRITES_ORIGINAL_OFFSET

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


Public Sub WriteRom()
 Dim Freedfile As Long
 
 Dim Index As Long
 Dim SubIndex As Long
 
 Dim RomPosition As Long

 Dim NamesPointerCount As Long

 If RomPath = "" Then
  Exit Sub
 End If
 
 Freedfile = FreeFile()

 ''''RomDump(&H17321) = 60

 
 'Save them spell names
 NamesPointerCount = pSpellNames ' 63940
 
 If SpellNamesInBounds = True Then
 RomPosition = pSpellNames '63940
 
 For Index = 0 To UBound(mSpellName())
 
  RomDump(RomPosition) = CByte(mSpellNameLength(Index))
  RomPosition = RomPosition + 1
   
  For SubIndex = 1 To mSpellNameLength(Index)
   RomDump(RomPosition) = AscB(Mid(mSpellName(Index), SubIndex, 1))
   RomPosition = RomPosition + 1
         
  Next SubIndex
 
  ''' Calculate the offset for the new pointers
  NamesPointerCount = NamesPointerCount + mSpellNameLength(Index) + 1
  
  If Index = 43 Then
   RomDump(ALLYNAMES_ORIGINAL_OFFSET + 1) = CByte(Fix(NamesPointerCount / 65536#))
   RomDump(ALLYNAMES_ORIGINAL_OFFSET + 2) = Fix((NamesPointerCount - Fix(NamesPointerCount / 65536#) * 65536#) / 256#)
   RomDump(ALLYNAMES_ORIGINAL_OFFSET + 3) = CByte(NamesPointerCount - Fix(Fix(NamesPointerCount / 256#) * 256#))
  End If
     
  If Index = 73 Then
   RomDump(ENEMYNAMES_ORIGINAL_OFFSET + 1) = CByte(Fix(NamesPointerCount / 65536#))
   RomDump(ENEMYNAMES_ORIGINAL_OFFSET + 2) = Fix((NamesPointerCount - Fix(NamesPointerCount / 65536#) * 65536#) / 256#)
   RomDump(ENEMYNAMES_ORIGINAL_OFFSET + 3) = CByte(NamesPointerCount - Fix(Fix(NamesPointerCount / 256#) * 256#))
  End If
     
 Next Index
 
 End If



 'Save them ITEM names
 NamesPointerCount = pItemNames '96622
 
 If ItemNamesInBounds = True Then
 RomPosition = pItemNames   '96622
 
 For Index = 0 To UBound(mItemName())
 
  RomDump(RomPosition) = CByte(mItemNameLength(Index))
  RomPosition = RomPosition + 1
   
  For SubIndex = 1 To mItemNameLength(Index)
   RomDump(RomPosition) = AscB(Mid(mItemName(Index), SubIndex, 1))
   RomPosition = RomPosition + 1
         
  Next SubIndex
 
  ''' Calculate the offset for the new pointers
  NamesPointerCount = NamesPointerCount + mItemNameLength(Index) + 1

  If Index = 127 Then
      
   RomDump(RomPosition) = 255
   RomPosition = RomPosition + 1
   
   NamesPointerCount = NamesPointerCount + 1
   
   RomDump(CLASSNAMES_ORIGINAL_OFFSET + 1) = CByte(Fix(NamesPointerCount / 65536#))
   RomDump(CLASSNAMES_ORIGINAL_OFFSET + 2) = Fix((NamesPointerCount - Fix(NamesPointerCount / 65536#) * 65536#) / 256#)
   RomDump(CLASSNAMES_ORIGINAL_OFFSET + 3) = CByte(NamesPointerCount - Fix(Fix(NamesPointerCount / 256#) * 256#))
  End If
     
 Next Index
 
 End If
 


 Open RomPath For Binary As #Freedfile
  Put #1, , RomDump

 Close #Freedfile

 MsgBox "ROM Saved~!", vbOKOnly
End Sub



