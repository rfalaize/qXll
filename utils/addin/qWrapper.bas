Attribute VB_Name = "qWrapper"
Option Explicit

'Wrappers
Public Function qwExecute(query As String, Optional synchronous As Boolean = False, Optional host As String = "localhost", Optional port As Integer = 5001) As Variant
    qwExecute = Application.Run("qExecute", query, synchronous, host, port)
End Function
Public Function qwQuery(query As String, Optional noHeaders As Boolean = False, Optional host As String = "localhost", Optional port As Integer = 5001) As Variant
    qwQuery = Application.Run("qQuery", query, noHeaders, host, port)
End Function
Public Function qwInsert(data As Variant, tableName As String, Optional createTable As Boolean = False, Optional keyedColumns As Integer = 0, Optional synchronous As Boolean = False, Optional host As String = "localhost", Optional port As Integer = 5001) As Variant
    qwInsert = Application.Run("qInsert", data, tableName, createTable, keyedColumns, synchronous, host, port)
End Function
