Attribute VB_Name = "lib_SupportFunctions"

'****************************************************************
'======== FUNCTIONS FOR SHEETS  =================================
'****************************************************************


Public Function concatRange(data As Range, Optional SEP As String = "") As String
    Dim ret As String
    Dim sep2 As String
    ret = ""
    sep2 = ""
    Dim vCell As Variant
    For Each vCell In data
        ret = ret & sep2 & vCell.value
        sep2 = SEP
    Next vCell

    concatRange = ret
End Function



Public Function SumProductSales(ByRef rngData As Range, ByRef arrData As Variant, ParamArray Conditions() As Variant) As Variant
'Returns array which equals to specified conditions
'only for digits
  Dim arrConditions As Variant
  arrConditions = Conditions
  'Set temp array accordingly to data
  Dim arrTemp As Variant
  arrTemp = rngData
  Dim Checked As Boolean
  ReDim Preserve arrConditions(LBound(arrConditions) To UBound(arrConditions) + 1)
  arrConditions(UBound(arrConditions)) = arrData
  
  'Add values to dictionary, we will use count number as Key
  Dim dictData As New Collection
  Set dictData = CreateDataDictionary(arrTemp)
  
  Dim i As Long
  'Check data to condition, if its true for all conditions, then return value
  Dim j As Long
  For j = LBound(arrTemp) To UBound(arrTemp)
    Checked = False
    For i = LBound(arrConditions) To UBound(arrConditions)
    
      Select Case IsArray(arrConditions(i))
        Case True
          Checked = DictCheckData(dictData, j, arrConditions(i)(j, 1))
        Case False
          Checked = DictCheckData(dictData, j, arrConditions(i))
      End Select
      
      Select Case Checked
        Case True
          Exit For
      End Select
    Next
  Next j
  
  Dim Key As Long
  'Return values from dictionary
  Dim arrExport As Variant
  ReDim arrExport(LBound(arrTemp) To UBound(arrTemp), 1 To 1)
  
  For Key = LBound(arrTemp) To dictData.Count
    arrExport(Key, 1) = dictData.item(CStr(Key))
  
  Next Key
  SumProductSales = arrExport
End Function



Public Function SheetExistence( _
  ByRef wbkActive As Workbook, _
  ByVal strSheetNameToFind As String, _
  ByVal blnSheetExists As Boolean) As Boolean
  
  Dim objSheet As Object
    For Each objSheet In wbkActive.Worksheets
      If strSheetNameToFind = objSheet.Name _
      And blnSheetExists = False Then
        SheetExistence = True
        Exit Function
      End If
    Next objSheet
    
End Function

Public Function addset_sht( _
  ByRef wbkActive As Workbook, _
  ByVal strSheetName As String) As Worksheet
  
  Dim blnSheetExists As Boolean
 
  blnSheetExists = SheetExistence( _
    wbkActive, _
    strSheetName, _
    False)
    
  With wbkActive
    Select Case blnSheetExists
      Case True
        Set addset_sht = .Sheets(strSheetName)
        Exit Function
    End Select
    ' If the sub goes here, then sheet is note exists
    'and we will create it
    Dim shtNew As Worksheet
    Set shtNew = .Worksheets.Add(After:=.Worksheets(.Worksheets.Count))
    shtNew.Name = strSheetName
    Set addset_sht = shtNew
  End With
    
End Function

Private Function isString(ByRef data As Variant) As Boolean

  Select Case VarType(data) = vbString
    Case True
      isString = True
    Case False
      isString = False
  End Select
  
End Function

Private Function DictCheckData(ByRef dictData As Collection, ByVal Key As Variant, ByVal DataToCheck As Variant) As Boolean
  Dim str As String
  str = dictData.item(CStr(Key))
  Select Case True
    Case DataToCheck = ">0"
      Select Case str > 0
        Case False
          DictRemoveKey dictData, CStr(Key)
          DictCheckData = True
          Exit Function
        Case True
          Exit Function
      End Select
    Case DataToCheck = "<0"
      Select Case str < 0
        Case False
          DictRemoveKey dictData, CStr(Key)
          DictCheckData = True
          Exit Function
        Case True
          Exit Function
      End Select
  End Select
  Select Case DataToCheck
    Case True
      Exit Function
    Case False
      DictRemoveKey dictData, CStr(Key)
      DictCheckData = True
      Exit Function
  End Select
  'in case if we here, then we use usual check
  Select Case str = DataToCheck
    Case False
      DictRemoveKey dictData, CStr(Key)
      DictCheckData = True
  End Select
  
End Function
Private Sub DictRemoveKey(ByRef dictData As Collection, ByVal Key As Variant)
  dictData.Remove CStr(Key)
  dictData.Add 0, CStr(Key)
End Sub


Private Function CreateDataDictionary(ByRef arrTemp As Variant) As Collection
'creates dictionary with data
  Dim i As Long
  Set CreateDataDictionary = New Collection
  
  For i = LBound(arrTemp) To UBound(arrTemp)
    CreateDataDictionary.Add arrTemp(i, 1), CStr(i)
  Next i

End Function



'****************************************************************
'======== PROGRESS BARS  =================================
'****************************************************************


'****************************************************************
'======== PROGRESS BARS  =================================
'****************************************************************

Public Sub ShowProgressBar()

' progress bar
ufProgress.LabelProgress.Width = 0
ufProgress.Show

End Sub

Public Sub CloseProgressBar()

' progress bar
Unload ufProgress

End Sub


Public Sub UpdateProgressBar(i As Variant, arr1 As Variant)
    Dim pctdone As Long
    pctdone = i / UBound(arr1)
    With ufProgress
        .LabelCaption.Caption = "Proc." & i & " of " & UBound(arr1)
        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
    End With
    DoEvents
End Sub

Public Sub UpdLabelProgressBar(i As Variant, q As Variant, Optional Label As String)
    Dim pctdone As Long
    pctdone = i / q
    With ufProgress
        .LabelCaption.Caption = Label & i & " of " & q
        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
    End With
    DoEvents
End Sub


'****************************************************************
'========= BASIC FUNCTIONS FOR VBA ==============
'****************************************************************

Function WhichColumn(vClear As Range) As Variant
'DEVELOPER: Anton Malofeev
'DESCRIPTION: Function to cut the cell address to column name, i.e.
' $A$2 -> A
Dim c As String
    c = vClear.Address
' cleaning 1 left symbol
    c = Right(c, Len(c) - 1)
' cleaning all symbols after and with second $
    c = Left(c, InStr(1, c, "$", vbTextCompare) - 1)
WhichColumn = c
End Function
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Function WhereInArray(arr1 As Variant, vFind As Variant) As Variant
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check where a value is in an array
Dim i As Long
For i = LBound(arr1) To UBound(arr1)
    If LCase(arr1(i)) = LCase(vFind) Then
        WhereInArray = i
        Exit Function
    End If
Next i
'if we get here, vFind was not in the array. Set to null
WhereInArray = Null
End Function

Function WhichRow(vClear As Variant) As Variant
'DEVELOPER: Anton Malofeev
'DESCRIPTION: Function to cut the cell address to column name, i.e.
' $A$2 -> 2
' cleaning 1 left symbol
    c = vClear.Address
    c = Right(c, Len(c) - 1)
' cleaning all symbols before and with second $
    c = Right(c, InStr(1, c, "$", vbTextCompare) - 1)
WhichRow = c
End Function
Function VerticalArr(vCell As Range, sht As String) As Variant
'DEVELOPER: Anton Malofeev
'DESCRIPTION: Function to cut the cell address to column name, i.e.
' $A$2 -> 2
    Dim lastColumn As Long
    lastColumn = Worksheets(sht).Range(vCell.Address).CurrentRegion.Columns.Count
    Dim a As String

    a = WhichColumn(vCell)
    COL = Letter2Number(a)
    If COL > 1 Then
        addcol = COL - 1
    End If
    b = WhichRow(vCell)
    c = Col_Letter(lastColumn + addcol)
    VerticalArr = Application.Transpose(Worksheets(sht).Range(a & b & ":" & c & b))
    VerticalArr = Application.Transpose(VerticalArr)
End Function
Function Letter2Number(vColname As String) As Variant
'PURPOSE: Convert a given letter into it's corresponding Numeric Reference
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim ColumnNumber As Long

'Convert To Column Number
   ColumnNumber = Range(vColname & 1).column

'Display Result
Letter2Number = ColumnNumber

End Function

Function WhereInArray2(arr1 As Variant, arr2 As Variant, vFind1 As Variant, vFind2 As Variant) As Variant
'DEVELOPER: Ryan Wells (wellsr.com), Anton Malofeev
'DESCRIPTION: Function to check where values is in arrays. Condition - equals position
  Dim i As Long
  For i = LBound(arr1) To UBound(arr1)
      If i = 710 Then
      k = 0
      End If
      If (arr1(i) = vFind1) And (arr2(i) = vFind2) = True Then
          WhereInArray2 = i
          Exit Function
      End If

  Next i

  'if we get here, vFind was not in the array. Set to null
  WhereInArray2 = Empty
End Function


Function ConvDigitString(data As Variant, StringOrDigit As String) As Variant

  Select Case StringOrDigit

  Case "String"
      ConvDigitString = CStr(data)
      Exit Function
  Case "Digit"
      On Error Resume Next
      ConvDigitString = CDbl(data)
      Exit Function
  End Select

End Function


'****************************************************************
'========= FUNCTIONS TO HANDLE WITH PRODUCT ARRAYS ==============
'****************************************************************

Function MakeArr(FirstR As Long, FirstC As Long, shtActive As Worksheet, _
        Optional FixRow As Long, Optional FixColumn As Long, _
        Optional strArrayOrRange As String, _
        Optional offsetFirstR As Long, Optional offsetFirstC As Long) As Variant
' FirstCell - first cell of range, i.e. cells(1,2)
' shtActive - working sheet with the array
' ExceptRowsQ - the POSITIVE number of rows which needs to delete from the array
' to get correct position (for example - title (-1))
' What is the height of the array?
' fix columns = true or false, are to fix range Current Region by column/row of First Cell
  Dim rngshtActFstCell As Range
  Dim ExceptRowsQ As Long
  Dim h As Long, w As Long
  Dim FR As Long, FC As Long
  With shtActive
    Set rngshtActFstCell = .Range(.Cells(FirstR, FirstC), .Cells(FirstR, FirstC))
  End With
        Select Case FixRow
        Case 0
          ExceptRowsQ = rngshtActFstCell.Row
          h = FirstR + rngshtActFstCell.CurrentRegion.Rows.Count - ExceptRowsQ - 1
        Case 1
          h = FirstR
        Case Is > 1
          h = FixRow
        End Select
      Select Case FixColumn
      Case 0
        w = FirstC - 1 + rngshtActFstCell.CurrentRegion.Columns.Count
      Case 1
        w = FirstC
      Case Is > 1
        w = FixColumn
      End Select
      'apply offsets
      FR = FirstR + offsetFirstR
      FC = FirstC + offsetFirstC
  With shtActive
  ' choose how it needs to show a result'
      Select Case strArrayOrRange
        Case "Range"
          Set MakeArr = .Range(.Cells(FR, FC), .Cells(FR + h - 1, w))
        Case Else 'Array
          MakeArr = .Range(.Cells(FR, FC), .Cells(FR + h - 1, w))
      End Select
  End With
  
End Function

Function GetPosID(UIID As String, SearchSht As Worksheet, _
        StartRow As Long) As Long
  Dim FirstR As Long, FirstC As Long

  With SearchSht

'The purpose is to define position of table
    Dim lastColumn As Long
    ' get last column position from a cell
    lastColumn = .Cells(1, 3)

    If lastColumn = 0 Then
      lastColumn = .Cells(StartRow, 1).CurrentRegion.Columns.Count
      FirstR = StartRow 'choose row for search Array'
    Else
      FirstR = 1
    End If
    FirstC = 1

    Dim ArrIDs As Variant
    ' make an array with ID of UI tables
    ArrIDs = .Range(.Cells(FirstR, FirstC), .Cells(FirstR, lastColumn))
    ArrIDs = Application.Transpose(ArrIDs)
    ArrIDs = Application.Transpose(ArrIDs)

    ' find column postion
    GetPosID = WhereInArray(ArrIDs, UIID)

  End With

End Function

Function GetArray(UIID As String, SearchSht As Worksheet, _
        StartRow As Long, Optional FixRow As Long, _
        Optional FixColumn As Long, Optional strArrayOrRange As String, _
        Optional offsetFirstR As Long, Optional offsetFirstC As Long) As Variant
' Searching position of UIID and making an arrray from it'
' StartRow is the row, from which an array will created
  Dim shtActive As Worksheet

  Set shtActive = ActiveSheet
  Dim cPos As Long
  cPos = GetPosID(UIID, SearchSht, StartRow)
  With shtActive
    ' choose how it needs to end'
    Select Case strArrayOrRange
    Case "Range"
      Set GetArray = MakeArr(StartRow, cPos, SearchSht, _
        FixRow, FixColumn, strArrayOrRange, offsetFirstR, offsetFirstC)
        Exit Function
    Case Else 'Array
      GetArray = MakeArr(StartRow, cPos, SearchSht, _
        FixRow, FixColumn, strArrayOrRange, offsetFirstR, offsetFirstC)
        Exit Function

    End Select

  End With

End Function

Public Function dictWFiles(ByVal sPath As String) As Dictionary
'Returns dictionary of files in specific directory
    Dim dictFiles     As Scripting.Dictionary
    Dim oFile       As Object
    Dim oFSO        As Object
    Dim oFolder     As Object
    Dim oFiles      As Object
    Set dictFiles = New Scripting.Dictionary
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.Files

    If oFiles.Count = 0 Then Exit Function

    For Each oFile In oFiles
        dictFiles.Add oFile.Name, oFile.Name
    Next

    Set dictWFiles = dictFiles

End Function



'simple
