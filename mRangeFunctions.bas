Attribute VB_Name = "mRangeFunctions"
Option Explicit


Public Function getLastRow( _
  ByRef shtWork As Worksheet, _
  Optional ByVal firstColumnIndexOffset As Long = 0 _
) As Long

    On Error GoTo ErrorHandler

    Dim checkNext As Boolean
    checkNext = True

    Dim LastRow As Long
    Dim firstRow As Long

    firstRow = 1
    Dim newLastRow As Long
    Dim i As Long
    Dim j As Long

    With shtWork
        Do While checkNext = True
        'check is it one row?
        
        For i = 1 To 4
            Dim isNextCellEmpty As Boolean
            isNextCellEmpty = checkIsEmpty(shtWork, firstRow + 1, i + firstColumnIndexOffset)
            If (Not (isNextCellEmpty)) Then
                GoTo checkColumns
            End If
        Next i
        LastRow = firstRow + 1
        GoTo checkEdningRows
      
checkColumns:
      'check three columns
      For i = 1 To 4
        newLastRow = .Cells(firstRow, i + firstColumnIndexOffset).End(xlDown).Row
        ' avoid limit problem
        If (newLastRow = 1048576) Then GoTo checkEdningRows
        
        Select Case True
          Case newLastRow > LastRow
            LastRow = newLastRow
        End Select
      Next i
      
checkEdningRows:
      'check three rows after lastrow
      For i = 1 To 4
        For j = 1 To 4
            Dim isCellEmpty As Boolean
            isCellEmpty = checkIsEmpty(shtWork, LastRow + j, i + firstColumnIndexOffset)
            If isCellEmpty Then
                checkNext = False
            Else
                firstRow = LastRow + j
                checkNext = True
                GoTo nextCheck
            End If
        Next j
      Next i
nextCheck:
    Loop
  End With

  getLastRow = LastRow
  
Exit Function
ErrorHandler:
    MsgBox prompt:="getLastRow" & Err.Description & " " & Err.Number
End Function
Public Function convertArrayToCollection( _
  ByRef arr As Variant, _
  ByVal exception As String _
) As Scripting.Dictionary
    On Error GoTo ErrorHandler
  Dim value As Variant
  Dim dic As Scripting.Dictionary
  Set dic = New Scripting.Dictionary
  Dim i As Long: i = LBound(arr)
  For Each value In arr
    If Not (value = exception) Then
        If Not (dic.Exists(value)) Then
          dic.Add CStr(value), i
        Else
          MsgBox prompt:="items are not unqiue" & " " & value & " " & i
          GoTo ErrorHandler
        End If
    End If
    i = i + 1
  Next value
  Set convertArrayToCollection = dic
Exit Function
ErrorHandler:
    MsgBox prompt:="convertArrayToCollection value" & Err.Description & " " & Err.Number
End Function
Public Function getLastColumn( _
  ByRef shtWork As Worksheet, _
  Optional ByVal firstColumnIndexOffset As Long = 0 _
) As Long
    On Error GoTo ErrorHandler
  Dim firstcell As Range
  Set firstcell = shtWork.Cells(1, 1 + firstColumnIndexOffset)
  
  getLastColumn = firstcell.CurrentRegion.Columns.Count
  
Exit Function
ErrorHandler:
    MsgBox prompt:="getLastColumn" & Err.Description & " " & Err.Number
End Function

Private Function checkIsEmpty( _
    ByRef sht As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long _
) As Boolean
    checkIsEmpty = LenB(sht.Cells(rowIndex, columnIndex).value) <= 0
End Function


'****************************************************************
'========= FUNCTIONS TO HANDLE  ARRAYS ==============
'****************************************************************
' Method to get full sized range
' row/column count can be used,
' if you need to fix size of range
Public Function getUsedRange( _
    ByVal rowIndex As Long, ByVal columnIndex As Long, _
    ByRef sheet As Worksheet, _
    Optional ByVal rowCount As Long, _
    Optional ByVal columnCount As Long _
) As Range
    On Error GoTo ErrorHandler
    ' get first cell
    Dim firstCellRange As Range
    Set firstCellRange = sheet.Cells(rowIndex, columnIndex)

    ' define last row
    Dim LastRow As Long
    If rowCount > 0 Then
        LastRow = rowCount
    Else
        LastRow = firstCellRange.CurrentRegion.Rows.Count
    End If
    
    ' define last right column
    Dim lastColumn As Long
    If columnCount > 0 Then
        lastColumn = columnCount
    Else
        lastColumn = firstCellRange.CurrentRegion.Columns.Count
    End If
    Dim finalRange As Range
    Set finalRange = firstCellRange.Resize(LastRow, lastColumn)

    Set getUsedRange = finalRange
    On Error GoTo 0
    Exit Function
ErrorHandler:
    MsgBox prompt:="getUsedRange" & Err.Description & " " & Err.Number
End Function

Public Function arrayFindIndex( _
    ByVal value As Variant, _
    ByRef arr As Variant, _
    Optional isExcelColumn As Boolean = False _
) As Long
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim el As Variant
    Dim arrayHasChildrens As Boolean
    
    For i = 1 To UBound(arr)
        If isExcelColumn Then
          el = arr(i, 1)
        Else
          el = arr(i)
        End If
        
        If IsArray(el) Then GoTo ErrorHandler
        
        If (el = value) Then
          arrayFindIndex = i
          Exit Function
        End If
    Next i
    ' item not found
    arrayFindIndex = -1

    On Error GoTo 0
    Exit Function
ErrorHandler:
    MsgBox prompt:="arrayFindIndex" & Err.Description & " " & Err.Number
End Function
