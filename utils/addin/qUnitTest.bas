Attribute VB_Name = "qUnitTest"
Option Explicit

'Unit tests
Public Sub UnitTest_qwExecute()
    Dim vo As Variant
    vo = qWrapper.qwExecute("n:50000;t:([]sym:n?`1;time:.z.p+til n;price:n?100.;size:n?1000)", False)
    Debug.Print vo(1)
End Sub
Public Sub UnitTest_qwInsert()
    Dim vi As Variant: ReDim vi(0 To 2, 0 To 1)
    vi(0, 0) = "sym": vi(0, 1) = "price"
    vi(1, 0) = "AAPL": vi(1, 1) = 145.33
    vi(2, 0) = "FB": vi(2, 1) = 77.64
    Dim vo As Variant
    vo = qWrapper.qwInsert(vi, "tInsert", True, 1)
    Debug.Print vo(1)
End Sub

Public Sub UnitTest_QueryAndDisplayInSelection()
    qUtils.QueryAndDisplayInSelection Selection, "t"
End Sub


