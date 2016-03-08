Attribute VB_Name = "GuyCounter"

Option Explicit

Public GuyBlocks(0 To 29) As Byte
Public GuyPointers(0 To 29) As Long

Public GuyNumber As Long

Public Function CountGuys() As Long

 Dim Index As Long
 Dim Count As Byte
 
 Index = pStats(0)  '&H1EE2F0
 
 Do While (RomDump(Index - 1) <> &HFF And RomDump(Index) <> &HFF) Or (RomDump(Index - 1) <> &HFE And RomDump(Index) <> &HFF)
 
  Index = Index + 1
  
  If RomDump(Index) = &HFF Or RomDump(Index) = &HFE Then
    Count = Count + 1
    Index = Index + 1
  End If
 
 Loop
 
 CountGuys = Count
 
 Call AssignBlocks
 
End Function


Private Sub AssignBlocks()
 
 Dim Index As Long
 Dim SubIndex As Long
 
 For Index = 0 To 29

   GuyPointers(Index) = CLng(CLng(RomDump(&H1EE271 + Index * 4)) * 65536 + CLng(RomDump(&H1EE271 + Index * 4 + 1)) * 256 + CLng(RomDump(&H1EE271 + Index * 4 + 2)))
   
 Next Index

 For Index = 0 To 29
  GuyBlocks(Index) = 0
 Next Index
  
 
 For Index = 0 To 28
 
   For SubIndex = GuyPointers(Index) To GuyPointers(Index + 1)
     If RomDump(SubIndex) = &HFF Or RomDump(SubIndex) = &HFE Then
     
     GuyBlocks(Index) = GuyBlocks(Index) + 1
     
     End If
   Next SubIndex
 
   GuyBlocks(Index + 1) = GuyBlocks(Index)
 Next Index

 'do the last one different dude  index = 29
 SubIndex = GuyPointers(29)
 
 Dim Continue As Boolean
 
 Continue = True
 
 Do While Continue
  
 If RomDump(SubIndex) = &HFF And RomDump(SubIndex + 1) = &HFF Then
  Continue = False
 ElseIf RomDump(SubIndex) = &HFE And RomDump(SubIndex + 1) = &HFF Then
  Continue = False
 End If
  
   If RomDump(SubIndex) = &HFF Or RomDump(SubIndex) = &HFE Then
     
    GuyBlocks(29) = GuyBlocks(Index) + 1
     
   End If
 
   SubIndex = SubIndex + 1
   
 Loop

 GuyBlocks(29) = GuyBlocks(29)
 'GuyBlocks(29) = GuyBlocks(Index)

End Sub


