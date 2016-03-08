Attribute VB_Name = "ControlFunctions"
Option Explicit

Public Sub SetCombo(ComboName As ComboBox, Value As Long)

 ComboName.ListIndex = 0
 Do While ComboName.ItemData(ComboName.ListIndex) <> Value And ComboName.ListIndex < ComboName.ListCount - 1
  ComboName.ListIndex = ComboName.ListIndex + 1
 Loop

 If ComboName.ItemData(ComboName.ListIndex) <> Value Then
  ComboName.ListIndex = -1
 End If

End Sub

Public Function ValidateInput(Value As TextBox, isLong As Boolean, Optional isChange As Boolean, Optional DeadRange As Byte) As Boolean

 If IsNumeric(Value.Text) = True Then
  If isLong = True Then
   
   If Value.Text <= 65535 - DeadRange And Value.Text >= 0 Then
    ValidateInput = True
   Else
    
    If isChange = False Then
     MsgBox "Invalid Input", vbOKOnly
     Value.SetFocus
    End If
    
    ValidateInput = False
   End If
  
  Else
   
   If Value.Text <= 255 - DeadRange And Value.Text >= 0 Then
    ValidateInput = True
   Else
    
    If isChange = False Then
     MsgBox "Invalid Input", vbOKOnly
     Value.SetFocus
    End If
    
    ValidateInput = False
   End If
  
  End If
 
 Else
 
  If isChange = False Then
   MsgBox "Invalid Input", vbOKOnly
   Value.SetFocus
  End If
  
  ValidateInput = False
 End If

End Function


Public Sub EnableTextboxes(EnableText As Boolean, FormName As Form)
 
 Dim Index As Long
 
 For Index = 0 To FormName.Count - 1
  
  If Left(FormName.Controls(Index).Name, 3) = "txt" Then
   FormName.Controls(Index).Enabled = EnableText
  End If
 
 Next Index

End Sub


' Looking for a damn way to return the charcode of a unicode character was
' getting far too damn tedious, so I wrote this. It certainly is slower
' than any native way of getting, but DAMN if no one ever seems
' to solve the type of problems I bump into...

Public Function ChCode(ByVal Character As String) As Byte

 Dim Index As Long
 
 Character = Left(Character, 1)

 Index = 0
 
 Do While Character <> Chr(Index) And Index < 255
 
  Index = Index + 1
 Loop
 
 ChCode = Index
 
End Function


Public Function RanNum(Min As Long, Max As Long) As Long
   'RanNum = CLng((Rnd * (Max - Min)) + Min)
   
   RanNum = CLng(Int((Rnd * (Max - Min + 1)) + Min))
End Function

Public Function FixedLength(ByVal StringText As String, ByVal Length As Integer) As String
  
  Do While Len(StringText) < Length
  
   StringText = " " & StringText
  
  Loop

  FixedLength = StringText
End Function

Public Function LeftLength(ByVal StringText As String, ByVal Length As Integer) As String
  
  Do While Len(StringText) < Length
  
   StringText = StringText & " "
  
  Loop

  LeftLength = StringText
End Function
