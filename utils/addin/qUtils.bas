Attribute VB_Name = "qUtils"
Option Explicit

'Utils
Public Sub QueryAndDisplayInRange(rng As Range, query As String, Optional noHeaders As Boolean = False, Optional host As String = "localhost", Optional port As Integer = 5001)
    Dim v As Variant: v = qWrapper.qwQuery(query, noHeaders, host, port)
    rng.Resize(UBound(v, 1) - LBound(v, 1) + 1, UBound(v, 2) - LBound(v, 2) + 1).Value = v
End Sub
Public Sub QueryAndDisplayInSelection(rng As Range, query As String, Optional noHeaders As Boolean = False, Optional host As String = "localhost", Optional port As Integer = 5001)
    QueryAndDisplayInRange Selection, query, noHeaders, host, port
End Sub
