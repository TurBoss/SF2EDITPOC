VERSION 5.00
Begin VB.Form Reader 
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   10005
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "Reader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  Dim Freedfile As Long
  Dim Texty As String
  Dim Feed As String
  
  Freedfile = FreeFile()
  
  Select Case Me.Caption
  Case "Users"
      
   Open App.Path & "/Data/Users.txt" For Input As #Freedfile
    Do While EOF(Freedfile) = False
     Line Input #Freedfile, Feed
     Texty = Texty & Feed & vbNewLine
    Loop
   Close #Freedfile
  Case "About"
  
    Open App.Path & "/Data/About.txt" For Input As #Freedfile
    Do While EOF(Freedfile) = False
     Line Input #Freedfile, Feed
     Texty = Texty & Feed & vbNewLine
    Loop
    Close #Freedfile
  Case "Developers"
  
    Open App.Path & "/Data/DevNotes.txt" For Input As #Freedfile
    Do While EOF(Freedfile) = False
     Line Input #Freedfile, Feed
     Texty = Texty & Feed & vbNewLine
    Loop
    Close #Freedfile
  End Select

  If Texty <> "" Then: Text.Text = Texty
End Sub

Private Sub Form_Resize()
  Text.Width = Abs(Me.Width - 300)
  Text.Height = Abs(Me.Height - 750)
End Sub
