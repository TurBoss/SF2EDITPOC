Attribute VB_Name = "modShops"
Option Explicit

Public ShopNames(0 To 30) As String
Public ShopSpot(0 To 30) As Long
Public ShopLength(0 To 30) As Long

Public StoreStart As Long
Public StoreEnd As Long


Public Sub Load_Shops()
  Dim Freedfile As Long
  Dim Index As Byte
  
  Index = 0
    
  StoreStart = SHOPITEMS_ORIGINAL_OFFSET
  StoreEnd = SHOPITEMS_ORIGINAL_OFFSET + SHOPITEMS_ORIGINAL_LENGTH - 1
    
  Freedfile = FreeFile()
    
  Open App.Path & "/Data/Stores.txt" For Input As #Freedfile
   
   For Index = 0 To 30
        
    Input #Freedfile, ShopNames(Index)
        
   Next

  Close #Freedfile
End Sub


Public Sub CalculateStoreSpots()

 Dim Index As Long
 Dim SubIndex As Long
 
 For Index = 0 To 30
 
  If Index = 0 Then
    ShopLength(Index) = RomDump(StoreStart)
    ShopSpot(Index) = StoreStart
    SubIndex = StoreStart + ShopLength(Index) + 1
  Else
    ShopLength(Index) = RomDump(SubIndex)
    ShopSpot(Index) = SubIndex
    
    SubIndex = SubIndex + ShopLength(Index) + 1
  End If
 Next
 
End Sub
