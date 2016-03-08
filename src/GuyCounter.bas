Attribute VB_Name = "GuyCounter"

Option Explicit

' Guyblocks is how many stat blocks a guy has
Public GuyBlocks() As Byte   ' (0 To 29)
Public GuyPointers() As Long

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
 
 ReDim GuyPointers(0)
 
 For Index = 0 To UBound(pStats())  ' 29

   If Index > 0 Then
    If pStats(Index) > pStats(Index - 1) Then
    
    ReDim Preserve GuyPointers(Index)
    GuyPointers(Index) = pStats(Index)
    'GuyPointers(Index) = CLng(CLng(RomDump(&H1EE271 + Index * 4)) * 65536 + CLng(RomDump(&H1EE271 + Index * 4 + 1)) * 256 + CLng(RomDump(&H1EE271 + Index * 4 + 2)))
    
    End If
   
   Else
    
    ReDim Preserve GuyPointers(Index)
    GuyPointers(Index) = pStats(Index)
    
   End If
   
 Next Index

 ReDim GuyBlocks(UBound(GuyPointers()))


 For Index = 0 To UBound(GuyBlocks())
  GuyBlocks(Index) = 0
 Next Index
  
 
 For Index = 0 To UBound(GuyPointers()) - 1
 
   For SubIndex = GuyPointers(Index) To GuyPointers(Index + 1)
     If RomDump(SubIndex) = &HFF Or RomDump(SubIndex) = &HFE Then
     
     GuyBlocks(Index) = GuyBlocks(Index) + 1
     
     End If
   Next SubIndex
 
   GuyBlocks(Index + 1) = GuyBlocks(Index)
 Next Index

 'do the last one different dude  index = 29
 SubIndex = GuyPointers(UBound(GuyPointers()))  ' GuyPointers(29)
 
 Dim Continue As Boolean
 
 Continue = True
 
 Do While Continue
  
 If RomDump(SubIndex) = &HFF And RomDump(SubIndex + 1) = &HFF Then
  Continue = False
 ElseIf RomDump(SubIndex) = &HFE And RomDump(SubIndex + 1) = &HFF Then
  Continue = False
 End If
  
   If RomDump(SubIndex) = &HFF Or RomDump(SubIndex) = &HFE Then
     
    GuyBlocks(UBound(GuyBlocks())) = GuyBlocks(Index) + 1
     
   End If
 
   SubIndex = SubIndex + 1
   
 Loop

 'GuyBlocks(29) = GuyBlocks(29)
 'GuyBlocks(29) = GuyBlocks(Index)

End Sub


