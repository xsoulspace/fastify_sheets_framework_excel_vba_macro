VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_value As Variant
Private m_rowIndex As Long
Private m_columnIndex As Long
Private m_headers As Variant
Private m_headerHeight As Long
Private m_sheetName As String

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub classInit( _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long, _
    ByVal sheetName As String, _
    Optional ByVal headerHeight As Long = 1 _
)
    On Error GoTo ErrorHandler

    m_rowIndex = rowIndex
    m_columnIndex = columnIndex
    m_sheetName = sheetName
    m_headerHeight = headerHeight
    On Error GoTo 0
    Exit Sub
ErrorHandler:
  MsgBox prompt:="classInit" & Err.Description & " " & Err.Number
End Sub

' Method to load full sized table
' based on x, y position
Public Sub loadTable()
    On Error GoTo ErrorHandler
    
    Dim tableRange As Range
    Set tableRange = mRangeFunctions.getUsedRange( _
        m_rowIndex + m_headerHeight, _
        m_columnIndex, m_sheet, Me.LastRow _
    )
    
    m_value = tableRange.Value2
    m_headers = getHeaders()

    On Error GoTo 0
    Exit Sub
ErrorHandler:
  MsgBox prompt:="loadTable" & Err.Description & " " & Err.Number
End Sub

' Method to write table
Public Sub writeTableToEnd( _
    ByRef values As Variant, _
    Optional ByVal offset As Long = 0 _
)
    On Error GoTo ErrorHandler
    
    Dim tableRange As Range
    Set tableRange = mRangeFunctions.getUsedRange( _
        Me.LastRow + 1 + offset, _
        m_columnIndex, Me.sheet, UBound(values), UBound(values, 2))
    
    tableRange.value = values
    
    On Error GoTo 0
    Exit Sub
ErrorHandler:
  MsgBox prompt:="writeTableToEnd" & Err.Description & " " & Err.Number
End Sub

Public Sub writeByHeaderName( _
    ByVal headerName As String, _
    ByRef values As Variant _
)
    On Error GoTo ErrorHandler
    Dim index As Long
    index = Me.getHeaderIndex(headerName)
    
    Dim columnRange As Range
    Set columnRange = mRangeFunctions.getUsedRange( _
        Me.rowIndex + 1, _
        index + m_columnIndex - 1, Me.sheet, UBound(values), 1)
    columnRange.Value2 = values
    
    On Error GoTo 0
    Exit Sub
ErrorHandler:
  MsgBox prompt:="writeByHeaderName" & Err.Description & " " & Err.Number
End Sub
Public Sub writeToRightNewLastColumn( _
  ByVal columnName As String, _
  ByRef values As Variant _
)
  On Error GoTo ErrorHandler
    Dim index As Long
    index = Me.lastColumn
    
    Dim columnRange As Range
    Set columnRange = mRangeFunctions.getUsedRange( _
        Me.rowIndex + 1, _
        index + m_columnIndex, Me.sheet, UBound(values), 1)
    columnRange.Value2 = values
    
  On Error GoTo 0
  Exit Sub
ErrorHandler:
    MsgBox prompt:="writeToRightNewLastColumn" & Err.Description & " " & Err.Number
End Sub

Public Function getHeaderIndex( _
  ByVal headerName As String _
) As Long
  On Error GoTo ErrorHandler
  Dim index As Long
  index = mRangeFunctions.arrayFindIndex(headerName, Me.headers)

  getHeaderIndex = index
  On Error GoTo 0
  Exit Function
ErrorHandler:
  getHeaderIndex = -1
End Function


Public Function getColumnValues( _
  ByVal columnName As String _
) As Variant
  On Error GoTo ErrorHandler
  ' find header index in headers
  Dim headerIndex As Long
  headerIndex = Me.getHeaderIndex(columnName)
  
  ' get values
  Dim values As Variant
  values = mRangeFunctions.getUsedRange( _
    m_rowIndex + m_headerHeight, _
    m_columnIndex + headerIndex - 1, _
    Me.sheet, _
    Me.LastRow - m_rowIndex - m_headerHeight + 1, 1 _
  )
  
  getColumnValues = values
  On Error GoTo 0
  Exit Function
ErrorHandler:
    MsgBox prompt:="getColumnValues" & Err.Description & " " & Err.Number
End Function


Public Function getColumnFirstValue( _
  ByVal columnName As String _
) As Variant
  On Error GoTo ErrorHandler
  ' find header index in headers
  Dim headerIndex As Long
  headerIndex = Me.getHeaderIndex(columnName)
  
  ' get values
  Dim values As Variant
  values = mRangeFunctions.getUsedRange( _
    m_rowIndex + m_headerHeight, _
    m_columnIndex + headerIndex - 1, _
    Me.sheet, _
    1, 1 _
  )
  
  getColumnFirstValue = values
  On Error GoTo 0
  Exit Function
ErrorHandler:
    MsgBox prompt:="getColumnFirstValue" & Err.Description & " " & Err.Number
End Function


Public Function getHeaders() As Variant
  On Error GoTo ErrorHandler
  Dim i As Long: i = 0
  Dim header As Variant: ReDim header(1 To Me.columnCount)

  Do While i <= m_headerHeight - 1

    Dim tableRange As Range
    Set tableRange = mRangeFunctions.getUsedRange( _
        m_rowIndex + i, _
        m_columnIndex, _
        Me.sheet, _
        1, Me.columnCount _
    )

    Dim values As Variant
    values = tableRange.Value2

    Dim value As Variant
    Dim j As Long: j = 1

    For Each value In values
        If LenB(value) > 0 Then
            header(j) = value
        End If
        j = j + 1
    Next value
    
    i = i + 1
  Loop
  getHeaders = header
  On Error GoTo 0
Exit Function
ErrorHandler:
  MsgBox prompt:="getHeaders" & Err.Description & " " & Err.Number
End Function

Public Property Get columnCount() As Long
  If IsEmpty(m_value) Then
    Dim firstcell As Range
    Set firstcell = Me.sheet.Cells(m_rowIndex, m_columnIndex)
    columnCount = firstcell.CurrentRegion.Columns.Count
  Else
    columnCount = UBound(m_value, 2)
  End If
End Property
Public Property Get rowIndex() As Long
    rowIndex = m_rowIndex
End Property
Public Property Get headers() As Variant
  headers = Me.getHeaders
End Property
Public Property Get headerHeight() As Long
  headerHeight = m_headerHeight
End Property
Public Property Get values() As Variant
  values = m_value
End Property

Public Property Get LastRow() As Variant
    LastRow = mRangeFunctions.getLastRow(Me.sheet, m_columnIndex - 1)
End Property
Public Property Get lastColumn() As Variant
    lastColumn = mRangeFunctions.getLastColumn(Me.sheet, m_columnIndex - 1)
End Property
Public Property Get sheet() As Worksheet
  Set sheet = Worksheets(m_sheetName)
End Property
