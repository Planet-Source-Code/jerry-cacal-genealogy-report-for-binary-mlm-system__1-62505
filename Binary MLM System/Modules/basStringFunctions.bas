Attribute VB_Name = "basStringFunctions"
Option Explicit

Public Function ParamValue(ParseCharacter As String, _
                               tString As Variant, _
                               Index As Integer) As String

  Dim CurrentPosition As Integer
  Dim ParseToPosition As Integer
  Dim CurrentToken As Integer
  Dim TempString As String
  TempString = Trim(tString) + ParseCharacter
  If Len(TempString) = 1 Then Exit Function
  CurrentPosition = 1
  CurrentToken = 1
  Do
    DoEvents
    ParseToPosition = InStr(CurrentPosition, TempString, _
                            ParseCharacter)
    If Index = CurrentToken Then
      ParamValue = Mid$(TempString, CurrentPosition, _
                        ParseToPosition - CurrentPosition)
      Exit Function
    End If
    CurrentToken = CurrentToken + 1
    CurrentPosition = ParseToPosition + 1
  Loop Until (CurrentPosition >= Len(TempString))
  
End Function

Public Function ParamCount(ParseCharacter As String, _
                               tString As Variant) As Integer
            
  Dim CurrentPosition As Integer
  Dim ParseToPosition As Integer
  Dim CurrentToken As Integer
  Dim TempString As String
  TempString = Trim(tString) + ParseCharacter
  If Len(TempString) = 1 Then Exit Function
  CurrentPosition = 1
  CurrentToken = 1
  Do
    DoEvents
    ParseToPosition = InStr(CurrentPosition, TempString, _
                            ParseCharacter)
    CurrentToken = CurrentToken + 1
    CurrentPosition = ParseToPosition + 1
  Loop Until (CurrentPosition >= Len(TempString))
  ParamCount = CurrentToken - 1
  If Right(Trim(tString), 1) = ParseCharacter Then
    ParamCount = ParamCount + 1
  End If
  
End Function
