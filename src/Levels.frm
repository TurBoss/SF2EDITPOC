VERSION 5.00
Begin VB.Form Levels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Levels"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   2550
   Begin VB.TextBox txtProLevel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtMaxLevelPro 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Promotion Level"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Max Level"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Max Level Promoted"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
      Begin VB.TextBox txtMaxLevel 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Levels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

 EnableTextboxes False, Me
 
 txtMaxLevel.Text = RomDump(38043)
 txtMaxLevelPro.Text = RomDump(38051)
 txtProLevel.Text = RomDump(134583)
 
 EnableTextboxes True, Me
 
End Sub

Private Sub txtMaxLevel_Change()
 Dim Result As Byte
 
 If ValidateInput(txtMaxLevel, False, True) = True Then
  Result = txtMaxLevel.Text
  RomDump(MAXLEVEL_ORIGINAL_OFFSET) = Result
 End If
End Sub

Private Sub txtMaxLevel_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtMaxLevel, False, False) = True Then
  Result = txtMaxLevel.Text
  RomDump(MAXLEVEL_ORIGINAL_OFFSET) = Result
 End If
End Sub

Private Sub txtMaxLevelPro_Change()
 Dim Result As Byte
 
 If ValidateInput(txtMaxLevelPro, False, True) = True Then
  Result = txtMaxLevelPro.Text
  RomDump(MAXLEVELUNPROMOTED_ORIGINAL_OFFSET) = Result
 End If
End Sub

Private Sub txtMaxLevelPro_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtMaxLevelPro, False, False) = True Then
  Result = txtMaxLevelPro.Text
  RomDump(MAXLEVELUNPROMOTED_ORIGINAL_OFFSET) = Result
 End If
End Sub

Private Sub txtProLevel_Change()
 Dim Result As Byte
 
 If ValidateInput(txtProLevel, False, True) = True Then
  Result = txtProLevel.Text
  RomDump(PROMOTIONLEVEL_ORIGINAL_OFFSET) = Result
  RomDump(135353) = Result ' Wiz note : Weapon item type equip check ... why change this ?
 End If
End Sub

Private Sub txtProLevel_LostFocus()
 Dim Result As Byte
 
 If ValidateInput(txtProLevel, False, False) = True Then
  Result = txtProLevel.Text
  RomDump(PROMOTIONLEVEL_ORIGINAL_OFFSET) = Result
  RomDump(135353) = Result ' Wiz note : Weapon item type equip check ... why change this ?
 End If
End Sub
