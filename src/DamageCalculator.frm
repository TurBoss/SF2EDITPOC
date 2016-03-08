VERSION 5.00
Begin VB.Form DamageCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Damage Calculator"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   4590
   Begin VB.Frame Frame5 
      Caption         =   "Damage"
      Height          =   855
      Left            =   2520
      TabIndex        =   10
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtDamage 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdChart 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   2580
      TabIndex        =   5
      Top             =   1245
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Critical"
      Height          =   855
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
      Begin VB.TextBox txtCritical 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Land Effect"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
      Begin VB.TextBox txtLE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "15"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Defense"
      Height          =   855
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   1095
      Begin VB.TextBox txtDef 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Attack"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
      Begin VB.TextBox txtAtk 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "DamageCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChart_Click()

 If ValidateInput(txtAtk, False) = True And ValidateInput(txtDef, False) = True And ValidateInput(txtLE, False) = True And ValidateInput(txtCritical, False) = True Then
 
 Else
  Exit Sub
 End If

 Dim Atk As Integer
 Dim Def As Integer
 Dim LdEf As Integer
 Dim Iterations As Long
 
 Dim MaxD As Long
 Dim MinD As Long
 
 Dim CritDam As Currency
 
 CritDam = txtCritical.Text
 Atk = txtAtk.Text
 Def = txtDef.Text
 LdEf = txtLE.Text
 Iterations = 500
 
 CritDam = CritDam / 100 + 1
 
 MaxD = Fix((Atk - Def) * CritDam) - Fix(Fix((Atk - Def) * CritDam) * (LdEf / 100))
 MinD = Fix(MaxD - Fix(MaxD * 0.25))
 
 If MinD < 1 Then: MinD = 1
 If MaxD < 1 Then: MaxD = 1
 
 txtDamage.Text = RanNum(MinD, MaxD) & "  [" & MinD & "-" & MaxD & "]"
 
 
End Sub
