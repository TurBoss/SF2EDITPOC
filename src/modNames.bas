Attribute VB_Name = "modNames"
 Option Explicit
  
 Public mItemName(159) As String
 Public mItemNameLength(159) As Long
 
 Public mSpellName(176) As String
 Public mSpellNameLength(176) As Long

 
 Public Function SpellNamesInBounds() As Boolean
 
  Dim Index As Long
  Dim Size As Long
  
  If isExpanded = True Then
   SpellNamesInBounds = True
   Exit Function
  End If
  
  
  For Index = 0 To UBound(mSpellName())
   
   Size = Size + Len(mSpellName(Index))

  Next Index

  If 63940 + Size + UBound(mSpellName()) + 1 < 65535 Then
    SpellNamesInBounds = True
  Else
    SpellNamesInBounds = False
  End If

 End Function
 
 
  Public Function ItemNamesInBounds() As Boolean
 
  Dim Index As Long
  Dim Size As Long
  
  
  If isExpanded = True Then
   ItemNamesInBounds = True
   Exit Function
  End If
    
    
  For Index = 0 To UBound(mItemName())
   
   Size = Size + Len(mItemName(Index))

  Next Index

  If 96622 + Size + UBound(mItemName()) + 1 < 98303 Then
    ItemNamesInBounds = True
  Else
    ItemNamesInBounds = False
  End If

 End Function
 
 
 Public Function GetSpellNameSize() As Long
 
  Dim Index As Long
  Dim Size As Long
  
  For Index = 0 To UBound(mSpellName())
   
   Size = Size + Len(mSpellName(Index))

  Next Index
  
  Size = Size + UBound(mSpellName()) + 1
  
  GetSpellNameSize = Size
 End Function

 Public Function GetItemNameSize() As Long
 
  Dim Index As Long
  Dim Size As Long
  
  For Index = 0 To UBound(mItemName())
   
   Size = Size + Len(mItemName(Index))

  Next Index
  
  Size = Size + UBound(mItemName()) + 1
  
  GetItemNameSize = Size
 End Function

 
 Public Sub LoadRomNames()

  Dim Index As Long
  Dim Offset As Long
 
  For Index = 0 To UBound(PersonName())
   PersonName(Index) = mSpellName(Index + 44)
  Next Index

  Offset = UBound(PersonName()) + 44 + 1

  For Index = 0 To UBound(MonsterName())
   MonsterName(Index) = mSpellName(Index + Offset)
  Next Index
   
  For Index = 0 To UBound(ItemName())
   ItemName(Index) = mItemName(Index)
  Next Index
  
  Offset = UBound(ItemName()) + 1
    
  For Index = 0 To UBound(ClassName())
   ClassName(Index) = mItemName(Index + Offset)
  Next Index
    
    
 End Sub








