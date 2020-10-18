Attribute VB_Name = "mcTableFactories"
Public Function cTableFactory( _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long, _
    ByVal sheetName As String, _
    Optional ByVal headerHeight As Long = 1 _
) As cTable
  On Error GoTo ErrorHandler

  Dim table As cTable
  Set table = New cTable

  ' Class initialization
  table.classInit rowIndex, columnIndex, sheetName, headerHeight

  ' return cTable instance
  Set cTableFactory = table
  
  On Error GoTo 0
  Exit Function
ErrorHandler:
  MsgBox prompt:="cTableFactory" & Err.Description & " " & Err.Number
End Function

Public Function settingsFactory() As cTableSettings
  On Error GoTo ErrorHandler

  ' Settings for cashbox
  Const rowIndex As Long = 1
  Const columnIndex As Long = 1
  Const sheetName As String = "Headers"

  Dim settingsBase As cTable

  Set settingsBase = cTableFactory( _
    rowIndex, _
    columnIndex, _
    sheetName _
  )

  Dim settings As cTableSettings
  Set settings = New cTableSettings
  Set settings.base = settingsBase
  ' return cTable instance
  Set settingsFactory = settings
  
  On Error GoTo 0
  Exit Function
ErrorHandler:
  MsgBox prompt:="settingsFactory" & Err.Description & " " & Err.Number
End Function

Public Function rawFactory() As cTable
  On Error GoTo ErrorHandler

  ' Settings for cashbox
  Const rowIndex As Long = 1
  Const columnIndex As Long = 1
  Dim sheetName As String
  sheetName = ActiveSheet.Name
  Dim raw As cTable

  Set raw = cTableFactory( _
    rowIndex, _
    columnIndex, _
    sheetName _
  )

  Set rawFactory = raw
  
  On Error GoTo 0
  Exit Function
ErrorHandler:
  MsgBox prompt:="rawFactory" & Err.Description & " " & Err.Number
End Function
Public Function stockFactory() As cTable
  On Error GoTo ErrorHandler

  ' Settings for cashbox
  Const rowIndex As Long = 1
  Const columnIndex As Long = 1
  Const sheetName As String = "stock"
  Dim stock As cTable

  Set stock = cTableFactory( _
    rowIndex, _
    columnIndex, _
    sheetName _
  )

  Set stockFactory = stock
  
  On Error GoTo 0
  Exit Function
ErrorHandler:
  MsgBox prompt:="stockFactory" & Err.Description & " " & Err.Number
End Function

