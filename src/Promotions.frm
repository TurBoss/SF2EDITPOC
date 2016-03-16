VERSION 5.00
Begin VB.Form Promotions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promotions"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   6495
   Begin VB.Frame Frame3 
      Caption         =   "Special Promotion Item"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   3015
      Begin VB.ComboBox cboSpecial 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   330
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Number of Basic Promotions"
      Height          =   855
      Left            =   3240
      TabIndex        =   7
      Top             =   5520
      Width           =   3135
      Begin VB.CommandButton cmdReallocate 
         Caption         =   "Reallocate"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox ClassNumber 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6255
      Begin VB.ListBox List 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   5775
      End
      Begin VB.ComboBox cboClass 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4800
         Width           =   2415
      End
      Begin VB.ComboBox cboPromotion 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   4845
         Width           =   855
      End
   End
End
Attribute VB_Name = "Promotions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Loading As Boolean

Private BasicClasses As Integer
Private SpecialClasses As Integer

Private PromotionItems As Integer


Private Sub cboClass_Click()
 
 If List.ListIndex = -1 Or Loading = True Then
  Exit Sub
 End If
 
 List.List(List.ListIndex) = cboClass.Text & " to " & cboPromotion.Text
 
 If List.ItemData(List.ListIndex) < PROMOTIONS_ORIGINAL_OFFSET + BasicClasses * 2 + 3 Then
  RomDump(List.ItemData(List.ListIndex)) = cboClass.ItemData(cboClass.ListIndex)
  RomDump(List.ItemData(List.ListIndex) + BasicClasses + 1) = cboPromotion.ItemData(cboPromotion.ListIndex)
 Else
  RomDump(List.ItemData(List.ListIndex)) = cboClass.ItemData(cboClass.ListIndex)
  RomDump(List.ItemData(List.ListIndex) + SpecialClasses + 1) = cboPromotion.ItemData(cboPromotion.ListIndex)
 End If
 
End Sub

Private Sub cboPromotion_Click()
 Call cboClass_Click
End Sub

Private Sub cboSpecial_Click()

 If Loading = False Then
  
  RomDump(PROMOTIONS_ORIGINAL_OFFSET + BasicClasses * 2 + SpecialClasses * 2 + 5 + List.ListIndex - BasicClasses) = cboSpecial.ItemData(cboSpecial.ListIndex)
   
 End If

End Sub

Private Sub cmdReallocate_Click()

 Dim Index As Long
 Dim SubIndex As Long
 Dim Counter As Integer
 
 Dim newBasicClasses As Integer
 Dim NotPromoted(0 To 16) As Integer
 Dim IsPromoted(0 To 16) As Integer
 
 
 If IsNumeric(ClassNumber.Text) = False Then
   MsgBox "Requires numeric", vbOKOnly
   Exit Sub
 End If
 
 newBasicClasses = CInt(ClassNumber.Text)
 
 If newBasicClasses > 17 Or newBasicClasses < 1 Then
   MsgBox "Value out of range", vbOKOnly
   Exit Sub
 End If


 
  SubIndex = RomDump(&H21046)
 
  Counter = 0
  
  For Index = PROMOTIONS_ORIGINAL_OFFSET + 1 To PROMOTIONS_ORIGINAL_OFFSET + SubIndex
   
   NotPromoted(Counter) = RomDump(Index)
   IsPromoted(Counter) = RomDump(Index + SubIndex + 1)
   
   Counter = Counter + 1
  Next
 
  For Index = PROMOTIONS_ORIGINAL_OFFSET + SubIndex * 2 + 3 To PROMOTIONS_ORIGINAL_OFFSET + SubIndex * 2 + 2 + 17 - SubIndex
  
   NotPromoted(Counter) = RomDump(Index)
   IsPromoted(Counter) = RomDump(Index + 17 - SubIndex + 1)
   
   Counter = Counter + 1
  Next

 
 Counter = 0
 
  
 '' For Index = &H21046 To &H21046 + 17 * 2 + 2
 
 Index = PROMOTIONS_ORIGINAL_OFFSET
   
 RomDump(Index) = newBasicClasses
 RomDump(Index + newBasicClasses + 1) = newBasicClasses
 
 For Index = PROMOTIONS_ORIGINAL_OFFSET + 1 To PROMOTIONS_ORIGINAL_OFFSET + newBasicClasses
   RomDump(Index) = NotPromoted(Counter)
   RomDump(Index + newBasicClasses + 1) = IsPromoted(Counter)
   
   Counter = Counter + 1
 Next Index
 
 SubIndex = PROMOTIONS_ORIGINAL_OFFSET + newBasicClasses * 2 + 2
 
 
 RomDump(SubIndex) = 17 - newBasicClasses
 RomDump(SubIndex + 17 - newBasicClasses + 1) = 17 - newBasicClasses
 
 For Index = SubIndex + 1 To SubIndex + 17 - newBasicClasses
   RomDump(Index) = NotPromoted(Counter)
   RomDump(Index + 17 - newBasicClasses + 1) = IsPromoted(Counter)
   
   Counter = Counter + 1
 Next Index
  
 Call GenerateList
 
 cboSpecial.Visible = False
  
 MsgBox "Reallocation Successful", vbOKOnly
   
  
End Sub

Private Sub GenerateList()
 Dim Index As Long
 
 Loading = True
 
 List.Clear
 
 BasicClasses = RomDump(PROMOTIONS_ORIGINAL_OFFSET)
 SpecialClasses = 17 - BasicClasses
 
 ClassNumber.Text = BasicClasses
 
 For Index = PROMOTIONS_ORIGINAL_OFFSET + 1 To PROMOTIONS_ORIGINAL_OFFSET + BasicClasses
 
  List.AddItem ClassName(RomDump(Index)) & " to " & ClassName(RomDump(Index + BasicClasses + 1))
  List.ItemData(List.ListCount - 1) = Index
  
 Next
 
 ' Special promotions
 
 If SpecialClasses > 0 Then
 
 For Index = PROMOTIONS_ORIGINAL_OFFSET + BasicClasses * 2 + 3 To PROMOTIONS_ORIGINAL_OFFSET + BasicClasses * 2 + 2 + SpecialClasses
 
  List.AddItem ClassName(RomDump(Index)) & " to " & ClassName(RomDump(Index + SpecialClasses + 1))
  List.ItemData(List.ListCount - 1) = Index
  
 Next
 
 End If
 
 Loading = False
End Sub

Private Sub Form_Load()

 Dim Index As Long
 
 Loading = True
 
 ' BasicClasses = RomDump(&H21046)
 ' SpecialClasses = 17 - BasicClasses
 
 BasicClasses = RomDump(PROMOTIONS_ORIGINAL_OFFSET)
 SpecialClasses = RomDump(PROMOTIONS_ORIGINAL_OFFSET + BasicClasses * 2 + 2)
 
 
 ClassNumber.Text = BasicClasses
 
 For Index = PROMOTIONS_ORIGINAL_OFFSET + 1 To PROMOTIONS_ORIGINAL_OFFSET + BasicClasses
 
  List.AddItem ClassName(RomDump(Index)) & " to " & ClassName(RomDump(Index + BasicClasses + 1))
  List.ItemData(List.ListCount - 1) = Index
  
 Next
 
 ' Special promotions
 
 If SpecialClasses > 0 Then
 
 For Index = PROMOTIONS_ORIGINAL_OFFSET + BasicClasses * 2 + 3 To PROMOTIONS_ORIGINAL_OFFSET + BasicClasses * 2 + 2 + SpecialClasses
 
  List.AddItem ClassName(RomDump(Index)) & " to " & ClassName(RomDump(Index + SpecialClasses + 1))
  List.ItemData(List.ListCount - 1) = Index
  
 Next
 
 End If
 
 
 For Index = 0 To 31
   
   cboClass.AddItem ClassName(Index)
   cboPromotion.AddItem ClassName(Index)
 
   cboClass.ItemData(Index) = Index
   cboPromotion.ItemData(Index) = Index
   
 Next
 
    
 For Index = 0 To UBound(ItemName())
  cboSpecial.AddItem ItemName(Index)
  cboSpecial.ItemData(Index) = Index
 Next
 
 Loading = False
 
End Sub

Private Sub List_Click()
 
 Loading = True
 
 cboClass.Enabled = True
 cboPromotion.Enabled = True
 
 If List.ListIndex <= BasicClasses - 1 Then
 
    SetCombo cboClass, CLng(RomDump(List.ItemData(List.ListIndex)))
    SetCombo cboPromotion, CLng(RomDump(List.ItemData(List.ListIndex) + BasicClasses + 1))
    
'   SetCombo cboClass, CLng(RomDump(&H21047 + List.ListIndex))
'   SetCombo cboPromotion, CLng(RomDump(&H21047 + List.ListIndex + 13))
 Else
 
    SetCombo cboClass, CLng(RomDump(List.ItemData(List.ListIndex)))
    SetCombo cboPromotion, CLng(RomDump(List.ItemData(List.ListIndex) + SpecialClasses + 1))
 
'   SetCombo cboClass, CLng(RomDump(&H21061 + List.ListIndex - 12))
'   SetCombo cboPromotion, CLng(RomDump(&H21061 + List.ListIndex - 12 + 6))
 End If
 
 If List.ListIndex >= BasicClasses Then
 
   PromotionItems = RomDump(PROMOTIONS_ORIGINAL_OFFSET + BasicClasses * 2 + SpecialClasses * 2 + 4)
   
   If List.ListIndex - BasicClasses < PromotionItems Then
   Loading = True
   
    SetCombo cboSpecial, CLng(RomDump(PROMOTIONS_ORIGINAL_OFFSET + BasicClasses * 2 + SpecialClasses * 2 + 5 + List.ListIndex - BasicClasses))
    cboSpecial.Visible = True
    
   Loading = False
   Else
    cboSpecial.Visible = False
   
   End If
 Else
   cboSpecial.Visible = False
 End If
 
 Loading = False
 
End Sub
