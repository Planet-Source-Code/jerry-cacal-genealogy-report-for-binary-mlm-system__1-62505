Attribute VB_Name = "basFileFunctions"
Option Explicit

Private mAdoCon As New ADODB.Connection
Private mAdoCom As New ADODB.Command
Private mAdoRec As New ADODB.Recordset

Public Function WriteTableToFile(sSQLString As String, sTextFileName As String, Optional bOverwriteTextFile As Boolean = True) As Boolean

  On Error GoTo OpenDatabaseErrorHandler
  
  Dim ResultSet$
  Dim ResultSetFileHandle&
  
  mAdoCon.Open MainDataBaseConnString
  mAdoCom.CommandType = adCmdText
  Set mAdoCom.ActiveConnection = mAdoCon
  mAdoCom.CommandText = sSQLString
  mAdoRec.CursorType = adOpenKeyset
  Set mAdoRec = mAdoCom.Execute
  
  If bOverwriteTextFile Then
    If Dir(sTextFileName) <> "" Then Kill sTextFileName
  End If
  
  ResultSetFileHandle = FreeFile
  Open sTextFileName For Output As #ResultSetFileHandle
  If mAdoRec.BOF Or mAdoRec.EOF Then
    ResultSet$ = vbNull
  Else
    ResultSet$ = mAdoRec.GetString(adClipString, , "|", vbCrLf, " ")
  End If
  Print #ResultSetFileHandle, ResultSet$
  Close #ResultSetFileHandle
  
  WriteTableToFile = True
  Set mAdoRec = Nothing
  Set mAdoCom = Nothing
  Set mAdoCon = Nothing
  Exit Function
  
OpenDatabaseErrorHandler:

  MsgBox Err.Description, vbOKOnly + vbCritical, "Write File Error"
  
  WriteTableToFile = False
  Set mAdoRec = Nothing
  Set mAdoCom = Nothing
  Set mAdoCon = Nothing
  
End Function

Public Function FileExists(FilePath_And_Name As String) As Boolean

  Dim fso
    
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  FileExists = False
  
  If Len(FilePath_And_Name) > 3 Then
    If InStr(1, FilePath_And_Name, ".") > 0 Then
      If fso.FileExists(FilePath_And_Name) Then
        FileExists = True
        Exit Function
      Else
        FileExists = False
        Exit Function
      End If
    Else
      FileExists = False
      Exit Function
    End If
  End If
  
End Function

Public Function FileToArray(TextFilePath As String) As Variant

  Dim tmp() As String
  Dim X As Long
  
  On Error GoTo FileErrHandler
  
  ReDim Preserve tmp(0)
  TextFilePath = Trim(TextFilePath)
    
  If FileExists(TextFilePath) Then
    Open TextFilePath For Input As #1
    X = 0
    Do While Not EOF(1)
      ReDim Preserve tmp(X + 1)
      Line Input #1, tmp(X)
      X = X + 1
    Loop
    
    If UBound(tmp) > 0 Then
      ReDim Preserve tmp(UBound(tmp) - 1)
    End If
    Close #1
    FileToArray = tmp
    Exit Function
  Else
    ReDim Preserve tmp(0)
    FileToArray = tmp
    Exit Function
  End If

FileErrHandler:
  ReDim Preserve tmp(0)
  FileToArray = tmp

End Function
