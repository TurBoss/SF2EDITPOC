Attribute VB_Name = "Disasm"
Option Explicit

Public Type DisasmFile
    id As String
    path As String
    bytes() As Byte
    modified As Boolean
End Type


Public DisasmConfFilePath As String
Public DisasmBasePath As String
Public DisasmFiles() As DisasmFile
Public AllyStatsFiles() As DisasmFile
Private Index As Byte



Public Sub LoadDisasm(FileAbsolutePath As String)
On Error GoTo OnError

    Call LoadDisasmConf(FileAbsolutePath)
    Call LoadDisasmFiles
    
    Main.mnuEdit.Enabled = True
    Main.mnuMisc.Enabled = True
    Main.mnuEditNames.Enabled = True

Exit Sub
OnError:
    MsgBox "Error while loading disasm. Index = " & Index, vbOKOnly
End Sub


Private Sub LoadDisasmConf(FileAbsolutePath As String)
On Error GoTo OnError
    Dim Freedfile As Long
    Dim logMessage As String
    
    DisasmConfFilePath = FileAbsolutePath
    DisasmBasePath = Left(DisasmConfFilePath, Len(DisasmConfFilePath) - Len(SF2EDITCONF_FILENAME))
    
    Freedfile = FreeFile()
    
    logMessage = "Loaded FilePaths :"

    Open DisasmConfFilePath For Input As #Freedfile
     Index = 0
     Do Until EOF(Freedfile)
      ReDim Preserve DisasmFiles(Index)
      Input #Freedfile, DisasmFiles(Index).id, DisasmFiles(Index).path
      logMessage = logMessage & vbNewLine & DisasmFiles(Index).id & " -> " & DisasmFiles(Index).path
      Index = Index + 1
     Loop
    Close #Freedfile

    MsgBox logMessage, vbOKOnly
Exit Sub
OnError:
    MsgBox "Error while loading disasm conf. Index = " & Index, vbOKOnly
End Sub


Private Sub LoadDisasmFiles()
On Error GoTo OnError
    
    Dim logMessage As String
        
    logMessage = "Loaded Files :"

    For Index = 0 To UBound(DisasmFiles)
     If DisasmFiles(Index).id <> "" Then
        If DisasmFiles(Index).id <> "AllyStats" Then
            Call LoadFile(DisasmFiles(Index))
            logMessage = logMessage & vbNewLine & DisasmFiles(Index).id & " -> " & UBound(DisasmFiles(Index).bytes) + 1 & " bytes"
        Else
            ' Load Ally Stats Files
            Call LoadAllyStatsFiles(DisasmFiles(Index))
        End If
     End If
    Next

    MsgBox logMessage, vbOKOnly
Exit Sub
OnError:
    MsgBox "Error while loading disasm files. Index = " & Index, vbOKOnly
End Sub


Private Sub LoadFile(DisasmFile As DisasmFile)
On Error GoTo OnError
     Dim Freedfile As Long
     Freedfile = FreeFile()
     Open DisasmBasePath & DisasmFile.path For Binary As #Freedfile
     ReDim DisasmFile.bytes(LOF(Freedfile) - 1)
     Get #Freedfile, , DisasmFile.bytes
     Close #Freedfile
Exit Sub
OnError:
    MsgBox "Error while loading disasm file. File : " & DisasmFile.path, vbOKOnly
End Sub


Private Sub LoadAllyStatsFiles(DisasmFile As DisasmFile)
On Error GoTo OnError
    Dim sFilename As String
    Dim logMessage As String
        
    logMessage = "Loaded Ally Stats Files :"
    
    sFilename = Dir(DisasmBasePath & DisasmFile.path & "*.bin")
    Index = 0
    Do While sFilename > ""
      ReDim Preserve AllyStatsFiles(Index)
      AllyStatsFiles(Index).id = sFilename
      AllyStatsFiles(Index).path = DisasmFile.path & sFilename
      Call LoadFile(AllyStatsFiles(Index))
      logMessage = logMessage & vbNewLine & AllyStatsFiles(Index).id & " -> " & UBound(AllyStatsFiles(Index).bytes) + 1 & " bytes"
      Index = Index + 1
      sFilename = Dir()
    Loop
        
    MsgBox logMessage, vbOKOnly
    
    Exit Sub
OnError:
    MsgBox "Error while loading ally stats file. File : " & DisasmFile.path, vbOKOnly
End Sub
