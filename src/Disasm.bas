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
Private logMessage As String



Public Sub LoadDisasm(FileAbsolutePath As String)
On Error GoTo OnError

    Call LoadDisasmConf(FileAbsolutePath)
    Call LoadDisasmFiles
    
    Main.mnuRomEditors.Enabled = False
    Main.mnuDisasmEditors.Enabled = True

Exit Sub
OnError:
    MsgBox "Error while loading disasm. Index = " & Index, vbOKOnly
End Sub


Private Sub LoadDisasmConf(FileAbsolutePath As String)
On Error GoTo OnError
    Dim Freedfile As Long
    
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
    Dim allyStatsLogMessage As String
    allyStatsLogMessage = "Loaded Ally Stats Files :"
    
    sFilename = Dir(DisasmBasePath & DisasmFile.path & "*.bin")
    Index = 0
    Do While sFilename > ""
      ReDim Preserve AllyStatsFiles(Index)
      AllyStatsFiles(Index).id = sFilename
      AllyStatsFiles(Index).path = DisasmFile.path & sFilename
      Call LoadFile(AllyStatsFiles(Index))
      allyStatsLogMessage = allyStatsLogMessage & vbNewLine & AllyStatsFiles(Index).id & " -> " & UBound(AllyStatsFiles(Index).bytes) + 1 & " bytes"
      Index = Index + 1
      sFilename = Dir()
    Loop
        
    MsgBox allyStatsLogMessage, vbOKOnly
    
    Exit Sub
OnError:
    MsgBox "Error while loading ally stats file. File : " & DisasmFile.path, vbOKOnly
End Sub

Public Function checkModifications() As Boolean
On Error GoTo OnError

    For Index = 0 To UBound(DisasmFiles)
     If DisasmFiles(Index).modified = True Then
        checkModifications = True
        Exit Function
     End If
    Next
    
    For Index = 0 To UBound(AllyStatsFiles)
     If AllyStatsFiles(Index).modified = True Then
        checkModifications = True
        Exit Function
     End If
    Next

    checkModifications = False

    Exit Function
OnError:
    MsgBox "Error while checking modifications. Index = " & Index, vbOKOnly
End Function

Public Sub WriteDisasmFiles()
On Error GoTo OnError
 Dim logMessage As String
 logMessage = "Updated disasm files :"
    For Index = 0 To UBound(DisasmFiles)
     If DisasmFiles(Index).id <> "AllyStats" And DisasmFiles(Index).modified = True Then
      Call WriteDisasmFile(DisasmFiles(Index))
     End If
    Next
    
    For Index = 0 To UBound(AllyStatsFiles)
     If AllyStatsFiles(Index).modified = True Then
        Call WriteDisasmFile(DisasmFiles(Index))
     End If
    Next
    MsgBox logMessage, vbOKOnly
    Exit Sub
OnError:
    MsgBox "Error while writing disasm files. Index = " & Index, vbOKOnly
End Sub

Private Sub WriteDisasmFile(DisasmFile As DisasmFile)
On Error GoTo OnError
 Dim Freedfile As Long
 Dim PreviousSize As Integer
 
 Open DisasmBasePath & DisasmFile.path For Binary As #Freedfile
 PreviousSize = LOF(Freedfile)
  Put #1, , DisasmFile.bytes
 Close #Freedfile
 logMessage = logMessage & vbNewLine & DisasmFile.id & " : " & PreviousSize & " bytes -> " & UBound(DisasmFile.bytes) + 1 & " bytes"

OnError:
    MsgBox "Error while writing disasm file. File : " & DisasmFile.id, vbOKOnly
End Sub
