Attribute VB_Name = "basMain"
Option Explicit

Public sTempFileDir As String
Public MainDataBaseConnString As String
Public TempDataBaseConnString As String
Public SQLStatementString As String

Sub Main()

  InitVars
  InitData
  frmMain.Show

End Sub

Public Sub InitVars()

  Dim MainDataBasePassword As String
  Dim MainDataBasePath     As String
  Dim MainDataBaseProvider As String
  Dim MainDataBaseSecurity As String
                                                     ' MainDataBaseProvider = "Microsoft.Jet.OLEDB.3.51" ' Version 3.51
  MainDataBaseProvider = "Microsoft.Jet.OLEDB.4.0"   ' Main Database Provider : use this for Access 2000
  MainDataBasePath = App.Path & "\Database\Data.mdb" ' Main MDB path & file
  MainDataBaseSecurity = "False"                     ' Main Database Security Info
  MainDataBasePassword = ""

  ' Main Database Connection String
  MainDataBaseConnString = "Provider=" & MainDataBaseProvider & ";" & _
                           "Data Source=" & Trim(MainDataBasePath) & ";" & _
                           "Persist Security Info=" & MainDataBaseSecurity & ";" & _
                           "Jet OLEDB:Database Password=" & Trim(MainDataBasePassword) & ";"

End Sub

Public Sub InitData()

  sTempFileDir = App.Path & "\membersinfo.txt"
  
  ' The mlmbinsys.txt file is located in the directory of the project.
  ' If you don't have it the system will automatically create it for you.
  ' Make sure the database exist and connected properly into the system.
  
  ' It is a text file of records in the database. I chose to store the records
  ' in a text file and processing from it making web programming much easier
  ' in the future.
  
  If Not WriteTableToFile("SELECT * FROM Members ORDER BY MemCode", sTempFileDir, True) Then
    MsgBox "Error writing to text file."
  End If

End Sub
