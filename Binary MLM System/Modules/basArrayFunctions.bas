Attribute VB_Name = "basArrayFunctions"
Option Explicit

Public Function SearchArray(ByRef ArrayData, ColumnIndex As Integer, ByVal SearchString As String, Optional SearchStart As Long = 0) As Long

  ' This module performs a sequential search in an array.
  ' To make a faster search change this module to
  ' binary search.
  
  Dim ctr As Long
  
  Dim StringFound As Boolean
  Dim String1 As String, String2 As String
  
  SearchArray = -1
  StringFound = False
  
  ctr = SearchStart
  
  Do
    DoEvents
    String1 = ParamValue("|", ArrayData(ctr), ColumnIndex)
    String2 = SearchString
    
    If String1 = String2 Then
      StringFound = True
      SearchArray = ctr
    End If
    
    ctr = ctr + 1
  Loop Until StringFound Or ctr >= UBound(ArrayData)

End Function

Public Function ArrayIndexValue(ByRef ArrayData, ColumnIndex As Integer, RowIndex As Long) As String

  If ColumnIndex > ParamCount("|", ArrayData) Or ColumnIndex < 0 Then
    ArrayIndexValue = ""
  Else
    ArrayIndexValue = ParamValue("|", ArrayData(RowIndex), ColumnIndex)
  End If

End Function
