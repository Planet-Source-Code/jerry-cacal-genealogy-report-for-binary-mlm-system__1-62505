Attribute VB_Name = "basBinFunctions"
Option Explicit

Dim DataArray As Variant
Dim TotalBinLeft As Long
Dim TotalBinRight As Long

Private Enum WingPosition
  wingLeft = 0
  wingRight = 1
  wingCenter = 2
End Enum

Private Sub BinaryStruct(ByVal MemberCode As String, _
                        Optional Wing As WingPosition = wingCenter)

  ' This is a recursive procedure to count the downlines of a member
  ' used only for counting the downlines
  
  Dim StartSearch As Long
  Dim CurrentPosition As Long
  Dim ArrayIndex As Integer
  Dim WingPos As WingPosition

  StartSearch = 0
  ArrayIndex = 8 ' UplineCode array index
  CurrentPosition = SearchArray(DataArray, ArrayIndex, MemberCode, StartSearch)

  While CurrentPosition <> -1
    DoEvents
    If Wing = wingLeft Then
      TotalBinLeft = TotalBinLeft + 1
      WingPos = wingLeft
    ElseIf Wing = wingRight Then
      TotalBinRight = TotalBinRight + 1
      WingPos = wingRight
    Else
      ArrayIndex = 9 ' BinaryPosition array index
      If ParamValue("|", DataArray(CurrentPosition), ArrayIndex) = "Left" Then
        TotalBinLeft = TotalBinLeft + 1
      End If
      If ParamValue("|", DataArray(CurrentPosition), ArrayIndex) = "Right" Then
        TotalBinRight = TotalBinRight + 1
      End If
      WingPos = IIf(ParamValue("|", DataArray(CurrentPosition), ArrayIndex) = "Left", wingLeft, _
                    IIf(ParamValue("|", DataArray(CurrentPosition), ArrayIndex) = "Right", wingRight, wingCenter))
    End If

    ArrayIndex = 3 ' MemberCode array index
    BinaryStruct ParamValue("|", DataArray(CurrentPosition), ArrayIndex), WingPos

    StartSearch = CurrentPosition + 1
    ArrayIndex = 8
    CurrentPosition = SearchArray(DataArray, ArrayIndex, MemberCode, StartSearch)
  Wend

End Sub

Public Sub BinaryInitialize()
  
  TotalBinLeft = 0
  TotalBinRight = 0
  
  DataArray = FileToArray(sTempFileDir)

End Sub

Public Function TotalLeft() As Long

  TotalLeft = TotalBinLeft

End Function

Public Function TotalRight() As Long

  TotalRight = TotalBinRight

End Function

Public Sub CountDownlines(ByVal sMemberID As String)

  TotalBinLeft = 0
  TotalBinRight = 0
  
  BinaryStruct sMemberID, wingCenter

End Sub

Public Sub FindNoUpline()

  Dim X As Long
  frmMain.Enabled = False
  frmWait.Show 0, frmMain
  For X = 0 To UBound(DataArray) - 1 Step 1
    DoEvents
    If SearchArray(DataArray, 3, ParamValue("|", DataArray(X), 8), 0) = -1 Then
      frmWait.Hide
      MsgBox "Upline code not found for" & vbCrLf & ParamValue("|", DataArray(X), 3) & " " & ParamValue("|", DataArray(X), 4) & vbCrLf & vbCrLf & "Upline code : " & ParamValue("|", DataArray(X), 8), vbOKOnly + vbInformation, "No Upline"
      frmWait.Show 0, frmMain
    End If
  Next X
  Unload frmWait
  frmMain.Enabled = True

End Sub

Private Function GetLevel(sMemberCodes As String) As String

  Dim tmp As String
  Dim X As Integer
  Dim SearchStart As Long
  Dim CurrentPosition As Long
  
  tmp = ""
  
  For X = 1 To ParamCount("|", sMemberCodes) Step 1
    DoEvents
    SearchStart = 0
    CurrentPosition = SearchArray(DataArray, 8, ParamValue("=", ParamValue("|", sMemberCodes, X), 2), SearchStart)
    While CurrentPosition <> -1
      DoEvents
      tmp = tmp + IIf(ParamCount("|", tmp) = 0, "", "|") + IIf(ParamValue("=", ParamValue("|", sMemberCodes, X), 1) = "C", IIf(ParamValue("|", DataArray(CurrentPosition), 9) = "Left", "L=", "R="), IIf(ParamValue("=", ParamValue("|", sMemberCodes, X), 1) = "L", "L=", "R=")) + ParamValue("|", DataArray(CurrentPosition), 3)
      SearchStart = CurrentPosition + 1
      CurrentPosition = SearchArray(DataArray, 8, ParamValue("=", ParamValue("|", sMemberCodes, X), 2), SearchStart)
    Wend
  Next X
  
  GetLevel = tmp

End Function

Public Function GetGenea(sMemberCode As String) As Variant

  Dim tmp() As String
  Dim Level As String
  Dim tmpLevel As String
  Dim X As Long
  
  ReDim tmp(0)
  
  tmp(0) = sMemberCode
  Level = sMemberCode
  tmpLevel = Level
  While ParamCount("|", Level) > 0
    DoEvents
    Level = GetLevel(tmpLevel)
    If ParamCount("|", Level) > 0 Then
      ReDim Preserve tmp(UBound(tmp) + 1)
      tmp(UBound(tmp)) = Level
    End If
    tmpLevel = Level
  Wend
  
  GetGenea = tmp
  
End Function

Public Function GetRecordInfo(sMemberCode As String) As String

  On Error Resume Next
  
  Dim UplineName As String, ReferrorName As String
  Dim BinPosition As String, MemberStatus As String
  
  ReferrorName = ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("=", sMemberCode, 2), 0)), 7), 0)), 4)
  UplineName = ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("=", sMemberCode, 2), 0)), 8), 0)), 4)
  BinPosition = ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("=", sMemberCode, 2), 0)), 9)
  MemberStatus = ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("=", sMemberCode, 2), 0)), 10)
  
  ReferrorName = IIf(Trim(ReferrorName) = "", "NONE", Trim(ReferrorName))
  UplineName = IIf(Trim(UplineName) = "", "NONE", Trim(UplineName))
  BinPosition = IIf(Trim(BinPosition) = "", "NONE", Trim(BinPosition))
  MemberStatus = IIf(Trim(MemberStatus) = "", "NONE", Trim(MemberStatus))
  
  GetRecordInfo = ParamValue("=", sMemberCode, 1) & "|" & _
                  ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("=", sMemberCode, 2), 0)), 2) & "|" & _
                  ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("=", sMemberCode, 2), 0)), 3) & "|" & _
                  ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("=", sMemberCode, 2), 0)), 4) & "|" & _
                  ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("=", sMemberCode, 2), 0)), 7) & "|" & _
                  ReferrorName & "|" & _
                  ParamValue("|", DataArray(SearchArray(DataArray, 3, ParamValue("=", sMemberCode, 2), 0)), 8) & "|" & _
                  UplineName & "|" & _
                  BinPosition & "|" & _
                  MemberStatus

End Function

Public Function MemberCodeExist(sMemberCode As String) As Boolean

  Dim CodeExists As Boolean
  Dim X As Long
  
  CodeExists = False
  
  For X = 0 To UBound(DataArray) Step 1
    If ParamValue("|", DataArray(X), 3) = ParamValue("=", sMemberCode, 2) Then
      MemberCodeExist = True
      Exit Function
    End If
  Next X
  
  MemberCodeExist = CodeExists

End Function

Public Function GetGenealogy(sMemberCode As String, lTotalLeft As Long, lTotalRight As Long, lLevels As Long) As Variant

  Dim Genea() As String
  Dim Levels As Variant
  Dim X As Long
  Dim Y As Integer
  Dim MemberCode As String
  Dim MemberInfo As String
  Dim LeftWing As Long, RightWing As Long
  
  LeftWing = 0
  RightWing = 0
  
  ReDim Preserve Genea(0)
  
  MemberCode = "C=" & sMemberCode
  If MemberCodeExist(MemberCode) Then
    Levels = GetGenea(MemberCode)
    For X = 0 To UBound(Levels) Step 1
      Genea(UBound(Genea)) = "LEVEL " & X
      ReDim Preserve Genea(UBound(Genea) + 1)
      For Y = 1 To ParamCount("|", Levels(X)) Step 1
        MemberInfo = GetRecordInfo(ParamValue("|", Levels(X), Y))
        Genea(UBound(Genea)) = MemberInfo
        ReDim Preserve Genea(UBound(Genea) + 1)
        Select Case ParamValue("|", MemberInfo, 1)
          Case "L"
            LeftWing = LeftWing + 1
          Case "R"
            RightWing = RightWing + 1
        End Select
      Next Y
    Next X
    lTotalLeft = LeftWing
    lTotalRight = RightWing
    lLevels = UBound(Levels)
    GetGenealogy = Genea
  Else
    ReDim Preserve Genea(0)
    lTotalLeft = 0
    lTotalRight = 0
    lLevels = 0
  End If

End Function
