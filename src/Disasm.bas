Attribute VB_Name = "Disasm"
Option Explicit

Public DisasmFilePaths() As String


Public Sub LoadDisasm(FileAbsolutePath As String)
On Error GoTo OnError

    Dim Freedfile As Long
    Dim Index As Byte
    Dim logMessage As String
    
    Freedfile = FreeFile()
    
    logMessage = "Loaded FilePaths :"

    Open FileAbsolutePath For Input As #Freedfile
     ReDim DisasmFilePaths(0)
     Index = 0
     Do While EOF(Freedfile) = False
      ReDim Preserve DisasmFilePaths(Index)
      Input #Freedfile, DisasmFilePaths(Index)
      logMessage = logMessage & vbNewLine & DisasmFilePaths(Index)
      Index = Index + 1
     Loop
    Close #Freedfile

    MsgBox logMessage, vbOKOnly

Exit Sub
OnError:
    MsgBox "Error while loading file.", vbOKOnly
End Sub


