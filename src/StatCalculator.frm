VERSION 5.00
Begin VB.Form StatCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stat Calculator"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   7830
   Begin VB.CommandButton cmdRange 
      Caption         =   "Range"
      Height          =   615
      Left            =   6360
      TabIndex        =   5
      Top             =   1245
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   2520
      TabIndex        =   11
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Iterations"
      Height          =   855
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   1095
      Begin VB.TextBox txtIterations 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "400"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   7575
      Begin VB.TextBox txtOutput 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   360
         Width           =   7095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Overwrite Base Stats - (HP, MP, ATK, DEF, AGI)"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   4575
      Begin VB.TextBox txtOverwrite 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdChart 
      Caption         =   "Average"
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Top             =   1245
      Width           =   1335
   End
   Begin VB.CommandButton cmdEstimate 
      Caption         =   "Single Level"
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   300
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Level"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
      Begin VB.TextBox txtLevel 
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
Attribute VB_Name = "StatCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Gro_Curve(0 To 5, 0 To 28) As Currency

'Private Gro_Linear(0 To 28) As Currency
'Private Gro_Late(0 To 28) As Currency
'Private Gro_Early(0 To 28) As Currency
'Private Gro_Middle(0 To 28) As Currency
'Private Gro_EarlyLate(0 To 28) As Currency

Private BaseStat(0 To 4) As Currency
Private oBaseStat(0 To 4) As Currency

Private GrowStat(0 To 4) As Currency
Private Stat(0 To 4) As Currency
Private CurveType(0 To 4) As Integer

Private Iterations As Long
Private iStat(0 To 4) As Currency

Private cStat(0 To 4, 0 To 39) As Currency

Private RangeLow(0 To 4, 0 To 39) As Currency
Private RangeHigh(0 To 4, 0 To 39) As Currency

Private cPosition As Long



Private Sub cmdChart_Click()

 Dim Index As Long
 
 'MsgBox "1"
 
 cPosition = txtOutput.SelStart
 
 'MsgBox "2"
 
 If CharacterStats.cboName.ListIndex < 0 Then: CharacterStats.cboName.ListIndex = 0
 
 'MsgBox "3"
 
 CharacterStats.Tag = 0
 
 'MsgBox "4"
 
 Me.SetFocus
 
' MsgBox "5"
 
 Me.Caption = "Stat Calculator" & " - " & CharacterStats.cboName.Text
  
 Dim OutputString As String

' MsgBox "6"

 Iterations = txtIterations.Text
  
 If Iterations > 1 Then
  OutputString = " LV      HP     MP     ATK     DEF     AGI" & vbNewLine & vbNewLine
 Else
  OutputString = " LV   HP   MP  ATK  DEF  AGI" & vbNewLine & vbNewLine
 End If
 
 'MsgBox "A"
  
 For Index = 0 To 39
  cStat(0, Index) = 0
  cStat(1, Index) = 0
  cStat(2, Index) = 0
  cStat(3, Index) = 0
  cStat(4, Index) = 0
 Next Index
  
 'MsgBox "7"
  
 For Index = 1 To Iterations
 
  Call SetStats
 
  Call GetLevel(40)
 
  DoEvents
 Next Index
 
'MsgBox "8"
 
 For Index = 0 To 39
  cStat(0, Index) = Round(cStat(0, Index) / Iterations, 2)
  cStat(1, Index) = Round(cStat(1, Index) / Iterations, 2)
  cStat(2, Index) = Round(cStat(2, Index) / Iterations, 2)
  cStat(3, Index) = Round(cStat(3, Index) / Iterations, 2)
  cStat(4, Index) = Round(cStat(4, Index) / Iterations, 2)
 Next Index
  
'MsgBox "9"

 For Index = 0 To 39
 
 If Iterations > 1 Then
 
 OutputString = OutputString & " " & FixedLength(CStr(Index + 1), 2) & "  " & FixedLength(Format(cStat(0, Index), "0.00"), 6) & " " & FixedLength(Format(cStat(1, Index), "0.00"), 6) & "  " & _
                FixedLength(Format(cStat(2, Index), "0.00"), 6) & "  " & FixedLength(Format(cStat(3, Index), "0.00"), 6) & "  " & _
                FixedLength(Format(cStat(4, Index), "0.00"), 6) & vbNewLine
 Else
 
 OutputString = OutputString & " " & FixedLength(CStr(Index + 1), 2) & "  " & FixedLength(cStat(0, Index), 3) & "  " & FixedLength(cStat(1, Index), 3) & "  " & _
                FixedLength(cStat(2, Index), 3) & "  " & FixedLength(cStat(3, Index), 3) & "  " & _
                FixedLength(cStat(4, Index), 3) & vbNewLine
 End If
                
 If (Index + 1) / 3 = Fix((Index + 1) / 3) Then: OutputString = OutputString & vbNewLine

 Next Index
                
'MsgBox "10"
                
 txtOutput.Text = OutputString
 
 txtOutput.SelStart = cPosition
End Sub



Private Sub cmdEstimate_click()

 Dim Index As Long
 
 If ValidateInput(txtLevel, False) = False Then
   Exit Sub
 End If

 If txtLevel.Text = "0" Then: MsgBox "Level 0 is not a valid level", vbOKOnly: Exit Sub
 
 If CharacterStats.cboName.ListIndex < 0 Then: CharacterStats.cboName.ListIndex = 0
 
 CharacterStats.Tag = 0
 
 Me.SetFocus
 
 Me.Caption = "Stat Calculator" & " - " & CharacterStats.cboName.Text
 
 Iterations = 1 'txtIterations.Text
 
 
 iStat(0) = 0
 iStat(1) = 0
 iStat(2) = 0
 iStat(3) = 0
 iStat(4) = 0
   
  
 For Index = 1 To Iterations
 
  Call SetStats
 
  Call GetLevel(CLng(txtLevel.Text))
 
  iStat(0) = iStat(0) + Stat(0)
  iStat(1) = iStat(1) + Stat(1)
  iStat(2) = iStat(2) + Stat(2)
  iStat(3) = iStat(3) + Stat(3)
  iStat(4) = iStat(4) + Stat(4)
 
  DoEvents
 Next Index
 
 iStat(0) = Round(iStat(0) / Iterations, 2)
 iStat(1) = Round(iStat(1) / Iterations, 2)
 iStat(2) = Round(iStat(2) / Iterations, 2)
 iStat(3) = Round(iStat(3) / Iterations, 2)
 iStat(4) = Round(iStat(4) / Iterations, 2)
 
 
 Dim OutputString As String
 

 
 If Iterations > 1 Then
 
 OutputString = " LV      HP     MP     ATK     DEF     AGI" & vbNewLine & vbNewLine
  
 OutputString = OutputString & " " & FixedLength(txtLevel.Text, 2) & "  " & FixedLength(Format(iStat(0), "0.00"), 6) & " " & FixedLength(Format(iStat(1), "0.00"), 6) & "  " & _
                FixedLength(Format(iStat(2), "0.00"), 6) & "  " & FixedLength(Format(iStat(3), "0.00"), 6) & "  " & _
                FixedLength(Format(iStat(4), "0.00"), 6)
 Else
 
 OutputString = " LV   HP   MP  ATK  DEF  AGI" & vbNewLine & vbNewLine
 
 OutputString = OutputString & " " & FixedLength(txtLevel.Text, 2) & "  " & FixedLength(iStat(0), 3) & "  " & FixedLength(iStat(1), 3) & "  " & _
                FixedLength(iStat(2), 3) & "  " & FixedLength(iStat(3), 3) & "  " & _
                FixedLength(iStat(4), 3)
 End If
                
 txtOutput.Text = OutputString
 
End Sub


Private Sub cmdRange_Click()

 Dim Index As Long
 
 cPosition = txtOutput.SelStart
 
 If CharacterStats.cboName.ListIndex < 0 Then: CharacterStats.cboName.ListIndex = 0
 
 CharacterStats.Tag = 0
 
 Me.SetFocus
 
 Me.Caption = "Stat Calculator" & " - " & CharacterStats.cboName.Text
  
 Dim OutputString As String
  
 Iterations = txtIterations.Text
  

  OutputString = " LV     HP       MP      ATK      DEF      AGI" & vbNewLine & vbNewLine

 
  
 For Index = 0 To 39
  RangeHigh(0, Index) = 0
  RangeHigh(1, Index) = 0
  RangeHigh(2, Index) = 0
  RangeHigh(3, Index) = 0
  RangeHigh(4, Index) = 0
  
  RangeLow(0, Index) = 1000
  RangeLow(1, Index) = 1000
  RangeLow(2, Index) = 1000
  RangeLow(3, Index) = 1000
  RangeLow(4, Index) = 1000
 Next Index
  
  
 For Index = 1 To Iterations
 
  Call SetStats
 
  Call GetLevel(40)
 
  DoEvents
 Next Index
 
 
 For Index = 0 To 39
  cStat(0, Index) = Round(cStat(0, Index) / Iterations, 2)
  cStat(1, Index) = Round(cStat(1, Index) / Iterations, 2)
  cStat(2, Index) = Round(cStat(2, Index) / Iterations, 2)
  cStat(3, Index) = Round(cStat(3, Index) / Iterations, 2)
  cStat(4, Index) = Round(cStat(4, Index) / Iterations, 2)
 Next Index
  
   
 For Index = 0 To 39
 


 OutputString = OutputString & " " & FixedLength(CStr(Index + 1), 2) & "  " & _
                FixedLength(RangeLow(0, Index) & "-", 4) & LeftLength(RangeHigh(0, Index), 3) & "  " & _
                FixedLength(RangeLow(1, Index) & "-", 4) & LeftLength(RangeHigh(1, Index), 3) & "  " & _
                FixedLength(RangeLow(2, Index) & "-", 4) & LeftLength(RangeHigh(2, Index), 3) & "  " & _
                FixedLength(RangeLow(3, Index) & "-", 4) & LeftLength(RangeHigh(3, Index), 3) & "  " & _
                FixedLength(RangeLow(4, Index) & "-", 4) & LeftLength(RangeHigh(4, Index), 3) & vbNewLine
 
                
 If (Index + 1) / 3 = Fix((Index + 1) / 3) Then: OutputString = OutputString & vbNewLine

 Next Index
                
 txtOutput.Text = OutputString
 
 txtOutput.SelStart = cPosition
 
End Sub

Private Sub Form_Load()
 Call Initialize_Growths
End Sub


Private Sub Initialize_Growths()

Dim strArray() As String
Dim inPoo As String
Dim Index As Long

inPoo = "313 352 352 352 352 313 352 352 352 352 352 313 352 352 352 352 352 313 352 352 352 352 352 313 352 352 352 352 352"

'inPoo = "3.13 3.52 3.52 3.52 3.52 3.13 3.52 3.52 3.52 3.52 3.52 3.13 3.52 3.52 3.52 3.52 3.52 3.13 3.52 3.52 3.52 3.52 3.52 3.13 3.52 3.52 3.52 3.52 3.52"

strArray = Split(inPoo, " ")

For Index = 0 To 28
 Gro_Curve(1, Index) = CCur(strArray(Index) / 100)
Next Index

inPoo = "234 273 273 273 313 273 313 313 313 313 313 313 352 352 352 352 352 352 352 391 391 391 391 352 430 391 430 391 469"
'inPoo = "2.34 2.73 2.73 2.73 3.13 2.73 3.13 3.13 3.13 3.13 3.13 3.13 3.52 3.52 3.52 3.52 3.52 3.52 3.52 3.91 3.91 3.91 3.91 3.52 4.30 3.91 4.30 3.91 4.69"

strArray = Split(inPoo, " ")

For Index = 0 To 28
 Gro_Curve(2, Index) = CCur(strArray(Index) / 100)
Next Index


inPoo = "430 430 391 430 391 352 430 352 391 391 391 313 352 352 352 352 352 313 313 313 313 352 273 273 313 273 273 313 234"

strArray = Split(inPoo, " ")

For Index = 0 To 28
 Gro_Curve(3, Index) = CCur(strArray(Index) / 100)
Next Index



inPoo = "234 273 313 313 313 273 352 352 391 352 391 352 430 430 430 430 430 352 391 352 391 352 352 273 313 313 313 273 273"

strArray = Split(inPoo, " ")

For Index = 0 To 28
 Gro_Curve(4, Index) = CCur(strArray(Index) / 100)
Next Index


inPoo = "469 469 430 430 391 313 352 313 313 313 273 273 234 273 273 273 234 273 273 313 313 313 352 313 391 430 430 469 508"

strArray = Split(inPoo, " ")

For Index = 0 To 28
 Gro_Curve(5, Index) = CCur(strArray(Index) / 100)
Next Index


End Sub


Private Sub SetStats()

 GrowStat(0) = CCur(CharacterStats.txtHpProjected.Text - CharacterStats.txtHP.Text)
 GrowStat(1) = CCur(CharacterStats.txtMpProjected.Text - CharacterStats.txtMP.Text)
 GrowStat(2) = CCur(CharacterStats.txtAtkProjected.Text - CharacterStats.txtAtk.Text)
 GrowStat(3) = CCur(CharacterStats.txtDefProjected.Text - CharacterStats.txtDef.Text)
 GrowStat(4) = CCur(CharacterStats.txtAgiProjected.Text - CharacterStats.txtAGI.Text)
 
 BaseStat(0) = CCur(CharacterStats.txtHP.Text)
 BaseStat(1) = CCur(CharacterStats.txtMP.Text)
 BaseStat(2) = CCur(CharacterStats.txtAtk.Text)
 BaseStat(3) = CCur(CharacterStats.txtDef.Text)
 BaseStat(4) = CCur(CharacterStats.txtAGI.Text)

 oBaseStat(0) = BaseStat(0)
 oBaseStat(1) = BaseStat(1)
 oBaseStat(2) = BaseStat(2)
 oBaseStat(3) = BaseStat(3)
 oBaseStat(4) = BaseStat(4)
 
 Call CheckforOverwrite

 Stat(0) = 0
 Stat(1) = 0
 Stat(2) = 0
 Stat(3) = 0
 Stat(4) = 0
 
 CurveType(0) = CInt(CharacterStats.cboHpGrowth.ListIndex)
 CurveType(1) = CInt(CharacterStats.cboMpGrowth.ListIndex)
 CurveType(2) = CInt(CharacterStats.cboAtkGrowth.ListIndex)
 CurveType(3) = CInt(CharacterStats.cboDefGrowth.ListIndex)
 CurveType(4) = CInt(CharacterStats.cboAgiGrowth.ListIndex)
 
 
End Sub


Private Function GetLevel(iLevel As Long)

 Dim Index As Long
 Dim Sindex As Long
 Dim PercentProgress(0 To 4) As Currency
 
 For Index = 1 To iLevel
 
  If Index = 1 Then
   Stat(0) = oBaseStat(0)
   Stat(1) = oBaseStat(1)
   Stat(2) = oBaseStat(2)
   Stat(3) = oBaseStat(3)
   Stat(4) = oBaseStat(4)
  Else
   
   If Index < 31 Then
   
    For Sindex = 0 To 4
     
     If CurveType(Sindex) > 0 Then
      Stat(Sindex) = Fix(Stat(Sindex) + GrowStat(Sindex) * (Gro_Curve(CurveType(Sindex), Index - 2) / 100) + RanNum(1, 99999) / 100000)
      PercentProgress(Sindex) = PercentProgress(Sindex) + Gro_Curve(CurveType(Sindex), Index - 2)
     End If
     
    Next Sindex
   
   Else
    
    For Sindex = 0 To 4
  
     If CurveType(Sindex) > 0 Then
      Stat(Sindex) = Stat(Sindex) + RanNum(1, 2)
     End If
          
    Next Sindex
        
    End If
    
    'Pity bonus
   
    For Sindex = 0 To 4
     If Stat(Sindex) < Fix(BaseStat(Sindex) + GrowStat(Sindex) * (PercentProgress(Sindex) / 100)) Then
     
      Stat(Sindex) = Stat(Sindex) + 1
     End If
     
    Next Sindex
   
  End If
  
  If Index < 41 Then
   cStat(0, Index - 1) = cStat(0, Index - 1) + Stat(0)
   cStat(1, Index - 1) = cStat(1, Index - 1) + Stat(1)
   cStat(2, Index - 1) = cStat(2, Index - 1) + Stat(2)
   cStat(3, Index - 1) = cStat(3, Index - 1) + Stat(3)
   cStat(4, Index - 1) = cStat(4, Index - 1) + Stat(4)
  End If
 
  If Index < 41 Then
   For Sindex = 0 To 4
   
   If Stat(Sindex) > RangeHigh(Sindex, Index - 1) Then
    RangeHigh(Sindex, Index - 1) = Stat(Sindex)
   End If
   
   If Stat(Sindex) < RangeLow(Sindex, Index - 1) Then
    RangeLow(Sindex, Index - 1) = Stat(Sindex)
   End If
    
   Next Sindex
  End If
 
 Next Index
  
 'Me.Caption = PercentProgress(0)

End Function


Private Sub CheckforOverwrite()

 Dim GetStuff() As String

 On Error GoTo GiveUp
 
 Do While Left(txtOverwrite.Text, 1) = " "
  txtOverwrite.Text = Right(txtOverwrite, Len(txtOverwrite.Text) - 1)
 Loop
 
 txtOverwrite.Text = Replace(txtOverwrite.Text, ",", " ")
 txtOverwrite.Text = Replace(txtOverwrite.Text, "  ", " ")
 txtOverwrite.Text = Replace(txtOverwrite.Text, "  ", " ")
 txtOverwrite.Text = Replace(txtOverwrite.Text, "  ", " ")
 txtOverwrite.Text = Replace(txtOverwrite.Text, "  ", " ")
 txtOverwrite.Text = Replace(txtOverwrite.Text, " ", ",")


 GetStuff = Split(txtOverwrite.Text, ",")

 If UBound(GetStuff()) = 4 Then

 oBaseStat(0) = CCur(GetStuff(0))
 oBaseStat(1) = CCur(GetStuff(1))
 oBaseStat(2) = CCur(GetStuff(2))
 oBaseStat(3) = CCur(GetStuff(3))
 oBaseStat(4) = CCur(GetStuff(4))
 
 ElseIf UBound(GetStuff()) = 5 Then
  
 oBaseStat(0) = CCur(GetStuff(1))
 oBaseStat(1) = CCur(GetStuff(2))
 oBaseStat(2) = CCur(GetStuff(3))
 oBaseStat(3) = CCur(GetStuff(4))
 oBaseStat(4) = CCur(GetStuff(5))
 
 End If

GiveUp:

End Sub


