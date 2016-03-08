Attribute VB_Name = "modItem"
' I'd have renamed this to modCode or modData, but I am far too lazy to.
' It won't affect the end user, thus should not waste the effort.

Option Explicit

Public ItemEquipCode() As Byte
Public ItemEquipBank() As Byte
' Public ItemEquipName(0 To 31) As String

Public ItemOffset() As Long
Public ItemOffsetName() As Long
Public ItemName() As String

Public AttributeName() As String
Public AttributeCode() As Byte

Public TypeName() As String
Public TypeCode() As Byte

Public SpellName() As String
Public SpellCode() As Byte

Public ItemCodeName() As String
Public ItemCode() As Byte

Public ClassName() As String  ' 0 To 31

'Public CharacterName(0 To 29) As String

Public PersonName() As String '0 To 29
Public PersonCode() As Byte

Public StatGrowthName() As String
Public StatGrowthCode() As Byte

Public MonsterName() As String

Public SpellTypeName() As String
Public SpellTypeCode() As Byte

Public MoveTypeName() As String
Public MoveTypeCode() As Byte

Public Sub Load_Codes()
  Dim Freedfile As Long
  Dim Index As Byte
  
  ' Gather equip codes
  
  Freedfile = FreeFile()

  Open App.Path & "/Data/ItemEquipCodes.txt" For Input As #Freedfile

   ReDim ItemEquipCode(0)
   ReDim ItemEquipBank(0)
   ReDim ClassName(0)
   
   Index = 0
   
   Do While EOF(Freedfile) = False   ' For Index = 0 To 31
    ReDim Preserve ItemEquipCode(Index)
    ReDim Preserve ItemEquipBank(Index)
    ReDim Preserve ClassName(Index)
    
    Input #Freedfile, ItemEquipCode(Index)
    Input #Freedfile, ItemEquipBank(Index)
    Input #Freedfile, ClassName(Index)
    
    Index = Index + 1
    'Input #Freedfile, ItemEquipName(Index)
    'ClassName(Index) = ItemEquipName(Index)
   Loop                              ' Next Index

  Close #Freedfile
  
  ' Get item infosSSss
  
  Index = 0
  
  Freedfile = FreeFile()
    
  ' The addresses for items are overwritten
  ' based on the location the pointer inside the rom.
  ' So the offsets loaded here are unused and redundant.
  
  Open App.Path & "/Data/ItemOffsets.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve ItemName(Index)
    ReDim Preserve ItemOffset(Index)
    ReDim Preserve ItemOffsetName(Index)
        
    Input #Freedfile, ItemName(Index)
    Input #Freedfile, ItemOffset(Index)
    Input #Freedfile, ItemOffsetName(Index)
         
     
    Index = Index + 1
   Loop

  Close #Freedfile
  
  ' Get att codes
  Index = 0
    
  Freedfile = FreeFile()
    
  Open App.Path & "/Data/AttributeCodes.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve AttributeCode(Index)
    ReDim Preserve AttributeName(Index)
    
    Input #Freedfile, AttributeCode(Index)
    Input #Freedfile, AttributeName(Index)
     
    Index = Index + 1
   Loop

  Close #Freedfile
  
  ''' Get type codes
  Index = 0
    
  Freedfile = FreeFile()
    
  Open App.Path & "/Data/ItemTypeCodes.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve TypeName(Index)
    ReDim Preserve TypeCode(Index)
    
    Input #Freedfile, TypeName(Index)
    Input #Freedfile, TypeCode(Index)
     
    Index = Index + 1
   Loop

  Close #Freedfile
  
  ''' Get spell codes. It'd be "more correct" to put this in a 'spell module',
  ''' but I'm doing items first, and I don't know if there even will be a spell
  ''' editor. Besides changing their names...
  
  Index = 0
    
  Freedfile = FreeFile()
    
  Open App.Path & "/Data/SpellCodes.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve SpellName(Index)
    ReDim Preserve SpellCode(Index)
    
    Input #Freedfile, SpellCode(Index)
    Input #Freedfile, SpellName(Index)
     
    Index = Index + 1
   Loop

  Close #Freedfile
  
  
  ' Get item codes...
  
  Index = 0
    
  Freedfile = FreeFile()
    
  Open App.Path & "/Data/ItemCodes.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve ItemCodeName(Index)
    ReDim Preserve ItemCode(Index)
    
    Input #Freedfile, ItemCode(Index)
    Input #Freedfile, ItemCodeName(Index)
     
    Index = Index + 1
   Loop

  Close #Freedfile
  
  ' Load character codes
  
'  Index = 0
'
'  Freedfile = FreeFile()
'
'  Open App.Path & "/Data/CharacterNames.txt" For Input As #Freedfile
'   Do While EOF(Freedfile) = False
'
'    Input #Freedfile, CharacterName(Index)
'
'    Index = Index + 1
'   Loop
'
'  Close #Freedfile
  
  
  ' Load person codes....
  
  Index = 0
    
  Freedfile = FreeFile()
    
  ReDim PersonName(0)
  ReDim PersonCode(0)
    
  Open App.Path & "/Data/PersonCodes.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve PersonName(Index)
    ReDim Preserve PersonCode(Index)
        
    Input #Freedfile, PersonName(Index)
    Input #Freedfile, PersonCode(Index)
     
    Index = Index + 1
   Loop

  Close #Freedfile
  
  
  ' Get stat growth codes...
  
  Index = 0
    
  Freedfile = FreeFile()
    
  Open App.Path & "/Data/StatGrowthCodes.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve StatGrowthName(Index)
    ReDim Preserve StatGrowthCode(Index)
    
    Input #Freedfile, StatGrowthCode(Index)
    Input #Freedfile, StatGrowthName(Index)
     
    Index = Index + 1
   Loop

  Close #Freedfile
  
  ' Get Monster ID's
  
  Index = 0
    
  Freedfile = FreeFile()
    
  Open App.Path & "/Data/MonsterCodes.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve MonsterName(Index)

    Input #Freedfile, MonsterName(Index)
     
    Index = Index + 1
   Loop

  Close #Freedfile
  
  ' Spell Type Codes's
  
  Index = 0
    
  Freedfile = FreeFile()
    
  Open App.Path & "/Data/SpellTypes.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve SpellTypeCode(Index)
    ReDim Preserve SpellTypeName(Index)
    
    Input #Freedfile, SpellTypeCode(Index)
    Input #Freedfile, SpellTypeName(Index)
    
    Index = Index + 1
   Loop

  Close #Freedfile
  
  ' Move Type Codes's
  
  Index = 0
    
  Freedfile = FreeFile()
    
  Open App.Path & "/Data/MoveCodes.txt" For Input As #Freedfile
   Do While EOF(Freedfile) = False
        
    ReDim Preserve MoveTypeCode(Index)
    ReDim Preserve MoveTypeName(Index)
    
    Input #Freedfile, MoveTypeCode(Index)
    Input #Freedfile, MoveTypeName(Index)
    
    Index = Index + 1
   Loop

  Close #Freedfile
End Sub

Public Sub Load_Path()

  Dim Freedfile As Long

On Error GoTo Failure

  Freedfile = FreeFile()

  Open App.Path & "/Data/Defaultpath.txt" For Input As #Freedfile

    Input #Freedfile, DefaultPath

  Close #Freedfile

 Exit Sub

Failure:

 Close #Freedfile

 DefaultPath = App.Path

End Sub



