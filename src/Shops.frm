VERSION 5.00
Begin VB.Form Shops 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shops"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   7950
   Begin VB.CommandButton cmdAddSlot 
      Caption         =   "Add Slot"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemoveSlot 
      Caption         =   "Remove Slot"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Inventory"
      Height          =   5535
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   7695
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add >>"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ListBox lstInventory 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4860
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
      Begin VB.ListBox lstItems 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4860
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Store Size"
      Height          =   735
      Left            =   6000
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      Begin VB.TextBox txtLength 
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      Begin VB.ComboBox cboName 
         Height          =   315
         ItemData        =   "Shops.frx":0000
         Left            =   120
         List            =   "Shops.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Shops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Address As Long
Private Length As Long
Private Sindex As Byte

Private Sub cboName_Click()
  Load_Shop cboName.ListIndex
    
  If cboName.ListIndex = cboName.ListCount - 1 Then
   cmdAddSlot.Enabled = False
   cmdRemoveSlot.Enabled = False
  Else
   cmdAddSlot.Enabled = True
   cmdRemoveSlot.Enabled = True
  End If
  
  txtLength.Text = ShopLength(cboName.ListIndex)
End Sub

Private Sub cmdAdd_Click()
 
 If lstInventory.ListIndex = -1 Or lstItems.ListIndex = -1 Then
  Exit Sub
 End If
 
 lstInventory.List(lstInventory.ListIndex) = lstItems.List(lstItems.ListIndex)
 lstInventory.ItemData(lstInventory.ListIndex) = lstItems.ItemData(lstItems.ListIndex)
 
 RomDump(Address + lstInventory.ListIndex) = lstItems.ItemData(lstItems.ListIndex)
 
End Sub

Private Sub cmdAddSlot_Click()
 
 Dim Index As Long
 Dim SubIndex As Long
 Dim CurrentStore As Long
   
 CurrentStore = cboName.ListIndex
 
 If ShopLength(30) > 0 Then
  ShopLength(CurrentStore) = ShopLength(CurrentStore) + 1

 
 For Index = 30 To CurrentStore + 1 Step -1
 
  If Index = 30 Then
    
    For SubIndex = ShopSpot(Index) + ShopLength(Index) - 1 To ShopSpot(Index) Step -1
     
     If SubIndex = ShopSpot(Index) Then
       ShopLength(30) = ShopLength(30) - 1
       RomDump(SubIndex + 1) = ShopLength(30)
     Else
      RomDump(SubIndex + 1) = RomDump(SubIndex)
     End If
     
    Next SubIndex
  Else
    
    For SubIndex = ShopSpot(Index) + ShopLength(Index) To ShopSpot(Index) Step -1

      RomDump(SubIndex + 1) = RomDump(SubIndex)
    
    Next SubIndex
    
  End If
   
  ShopSpot(Index) = ShopSpot(Index) + 1
 Next Index
 
 RomDump(ShopSpot(CurrentStore)) = ShopLength(CurrentStore)
 RomDump(ShopSpot(CurrentStore) + ShopLength(CurrentStore)) = 0
 
 End If
  
 Call cboName_Click
 
End Sub

Private Sub cmdRemoveSlot_Click()
 
 Dim Index As Long
 Dim SubIndex As Long
 Dim CurrentStore As Long
   
 CurrentStore = cboName.ListIndex
 
 If ShopLength(CurrentStore) > 1 Then
  ShopLength(CurrentStore) = ShopLength(CurrentStore) - 1

 
 For Index = CurrentStore + 1 To 30
 
  If Index = 30 Then
    
    For SubIndex = ShopSpot(Index) To ShopSpot(Index) + ShopLength(Index)
     
     If SubIndex = ShopSpot(Index) Then
       ShopLength(30) = ShopLength(30) + 1
       RomDump(SubIndex - 1) = ShopLength(30)
     Else
      RomDump(SubIndex - 1) = RomDump(SubIndex)
     End If
     
    Next SubIndex
  Else
    
    For SubIndex = ShopSpot(Index) To ShopSpot(Index) + ShopLength(Index)

      RomDump(SubIndex - 1) = RomDump(SubIndex)
    
    Next SubIndex
    
  End If
   
  ShopSpot(Index) = ShopSpot(Index) - 1
 Next Index
 
 RomDump(ShopSpot(CurrentStore)) = ShopLength(CurrentStore)
 
 End If
  
 Call cboName_Click
End Sub

Private Sub Form_Load()

 Dim Index As Long
 
 For Index = 0 To UBound(ItemCodeName())
  
  lstItems.AddItem ItemCodeName(Index)
  lstItems.ItemData(Index) = ItemCode(Index)

 Next
  
 For Index = 0 To UBound(ShopNames())
  cboName.AddItem ShopNames(Index)
 Next Index
End Sub

Private Sub Load_Shop(ShopIndex As Long)

 Dim Index As Long
 Dim SubIndex As Long
 
 lstInventory.Clear
 
 Address = ShopSpot(ShopIndex) + 1
 Length = ShopLength(ShopIndex)
 Sindex = ShopIndex
 
 For Index = ShopSpot(ShopIndex) + 1 To ShopSpot(ShopIndex) + ShopLength(ShopIndex)
   
   SubIndex = 0
   
   Do While SubIndex <> RomDump(Index) And SubIndex <= &H7F
    SubIndex = SubIndex + 1
   Loop
    
   lstInventory.AddItem ItemCodeName(SubIndex)
   lstInventory.ItemData(lstInventory.ListCount - 1) = ItemCode(SubIndex)
      
 Next
End Sub


