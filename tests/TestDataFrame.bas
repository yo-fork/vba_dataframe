Attribute VB_Name = "TestDataFrame"
Option Explicit

' ============================================================
'  VBA DataFrame - Test Suite
' ============================================================
'  Run TestAll from Immediate Window or F5
'  Results: Immediate Window (Ctrl+G)
' ============================================================

Private passCount As Long
Private failCount As Long

Sub TestAll()
    passCount = 0: failCount = 0
    Debug.Print String$(50, "=")
    Debug.Print "  DataFrame Test Suite"
    Debug.Print String$(50, "=")
    
    Test_Create
    Test_FromArray
    Test_Properties
    Test_Head_Tail
    Test_Sel
    Test_Where
    Test_OrderBy
    Test_AddCol
    Test_RemoveCol
    Test_RenameCol
    Test_Distinct
    Test_Slice
    Test_Sum_Mean
    Test_MinMax
    Test_Describe
    Test_GroupBy
    Test_Join
    Test_VStack
    Test_EmptyDF
    
    Debug.Print String$(50, "=")
    Debug.Print "  PASS: " & passCount & "  FAIL: " & failCount
    Debug.Print String$(50, "=")
End Sub

Private Sub Assert(ByVal testName As String, ByVal condition As Boolean, _
                   Optional ByVal msg As String = "")
    If condition Then
        passCount = passCount + 1
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL: " & testName & IIf(msg <> "", " - " & msg, "")
    End If
End Sub

' ============================================================
Sub Test_Create()
    Debug.Print ">> Test_Create"
    Dim df As DataFrame
    Set df = DFrame.Create( _
        Array("A", "B", "C"), _
        Array(1, "x", True), _
        Array(2, "y", False))
    Assert "Create.RowCount", df.RowCount = 2
    Assert "Create.ColCount", df.ColCount = 3
    Assert "Create.Value(1,A)", df.Value(1, "A") = 1
    Assert "Create.Value(2,B)", df.Value(2, "B") = "y"
End Sub

Sub Test_FromArray()
    Debug.Print ">> Test_FromArray"
    Dim data(1 To 3, 1 To 2) As Variant
    data(1, 1) = 10: data(1, 2) = "a"
    data(2, 1) = 20: data(2, 2) = "b"
    data(3, 1) = 30: data(3, 2) = "c"
    Dim df As DataFrame
    Set df = DFrame.FromArray(data, Array("Num", "Letter"))
    Assert "FromArray.RowCount", df.RowCount = 3
    Assert "FromArray.Value(3,Num)", df.Value(3, "Num") = 30
End Sub

Sub Test_Properties()
    Debug.Print ">> Test_Properties"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("X", "Y"), Array(1, 2), Array(3, 4), Array(5, 6))
    Assert "Shape", df.Shape = "3 rows x 2 cols"
    
    Dim cols As Variant: cols = df.Columns
    Assert "Columns(1)", cols(1) = "X"
    Assert "Columns(2)", cols(2) = "Y"
    
    Dim c As Variant: c = df.Col("X")
    Assert "Col.count", UBound(c) = 3
    Assert "Col(2)", c(2) = 3
    
    Dim rw As Variant: rw = df.Row(2)
    Assert "Row(2,1)", rw(1) = 3
    Assert "Row(2,2)", rw(2) = 4
End Sub

Sub Test_Head_Tail()
    Debug.Print ">> Test_Head_Tail"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("V"), Array(1), Array(2), Array(3), Array(4), Array(5))
    
    Dim h As DataFrame: Set h = df.Head(3)
    Assert "Head.RowCount", h.RowCount = 3
    Assert "Head.Last", h.Value(3, "V") = 3
    
    Dim t As DataFrame: Set t = df.Tail(2)
    Assert "Tail.RowCount", t.RowCount = 2
    Assert "Tail.First", t.Value(1, "V") = 4
End Sub

Sub Test_Sel()
    Debug.Print ">> Test_Sel"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("A", "B", "C"), Array(1, 2, 3), Array(4, 5, 6))
    Dim s As DataFrame: Set s = df.Sel("C", "A")
    Assert "Sel.ColCount", s.ColCount = 2
    Dim cols As Variant: cols = s.Columns
    Assert "Sel.Col1", cols(1) = "C"
    Assert "Sel.Col2", cols(2) = "A"
    Assert "Sel.Value", s.Value(1, "C") = 3
End Sub

Sub Test_Where()
    Debug.Print ">> Test_Where"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("N", "V"), _
        Array("a", 10), Array("b", 20), Array("c", 30), Array("d", 5))
    
    Assert "Where.GT", df.Where("V", ">", 10).RowCount = 2
    Assert "Where.GTE", df.Where("V", ">=", 10).RowCount = 3
    Assert "Where.EQ", df.Where("V", "=", 20).RowCount = 1
    Assert "Where.NE", df.Where("V", "<>", 20).RowCount = 3
    Assert "Where.LT", df.Where("V", "<", 20).RowCount = 2
    Assert "Where.Like", df.Where("N", "Like", "[ab]").RowCount = 2
    Assert "Where.In", df.Where("N", "In", Array("a", "c")).RowCount = 2
End Sub

Sub Test_OrderBy()
    Debug.Print ">> Test_OrderBy"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("N", "V"), _
        Array("c", 30), Array("a", 10), Array("b", 20))
    
    Dim asc_ As DataFrame: Set asc_ = df.OrderBy("V", True)
    Assert "OrderByAsc.First", asc_.Value(1, "V") = 10
    Assert "OrderByAsc.Last", asc_.Value(3, "V") = 30
    
    Dim desc_ As DataFrame: Set desc_ = df.OrderBy("V", False)
    Assert "OrderByDesc.First", desc_.Value(1, "V") = 30
End Sub

Sub Test_AddCol()
    Debug.Print ">> Test_AddCol"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("A"), Array(1), Array(2))
    
    ' Scalar
    Dim df2 As DataFrame: Set df2 = df.AddCol("B", "x")
    Assert "AddCol.ColCount", df2.ColCount = 2
    Assert "AddCol.Scalar", df2.Value(1, "B") = "x"
    
    ' Array
    Dim df3 As DataFrame: Set df3 = df.AddCol("C", Array(10, 20))
    Assert "AddCol.Array", df3.Value(2, "C") = 20
End Sub

Sub Test_RemoveCol()
    Debug.Print ">> Test_RemoveCol"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("A", "B", "C"), Array(1, 2, 3))
    Dim df2 As DataFrame: Set df2 = df.RemoveCol("B")
    Assert "RemoveCol.ColCount", df2.ColCount = 2
    Dim cols As Variant: cols = df2.Columns
    Assert "RemoveCol.Cols", cols(1) = "A" And cols(2) = "C"
End Sub

Sub Test_RenameCol()
    Debug.Print ">> Test_RenameCol"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("A", "B"), Array(1, 2))
    Dim df2 As DataFrame: Set df2 = df.RenameCol("A", "Alpha")
    Dim cols As Variant: cols = df2.Columns
    Assert "RenameCol", cols(1) = "Alpha"
    Assert "RenameCol.Value", df2.Value(1, "Alpha") = 1
End Sub

Sub Test_Distinct()
    Debug.Print ">> Test_Distinct"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("A", "B"), _
        Array(1, "x"), Array(2, "y"), Array(1, "x"), Array(3, "z"))
    Assert "Distinct.All", df.Distinct().RowCount = 3
    Assert "Distinct.Col", df.Distinct("A").RowCount = 3
End Sub

Sub Test_Slice()
    Debug.Print ">> Test_Slice"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("V"), Array(10), Array(20), Array(30), Array(40))
    Dim s As DataFrame: Set s = df.Slice(2, 3)
    Assert "Slice.RowCount", s.RowCount = 2
    Assert "Slice.First", s.Value(1, "V") = 20
    Assert "Slice.Last", s.Value(2, "V") = 30
End Sub

Sub Test_Sum_Mean()
    Debug.Print ">> Test_Sum_Mean"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("V"), Array(10), Array(20), Array(30))
    Assert "Sum", df.Sum("V") = 60
    Assert "Mean", df.Mean("V") = 20
End Sub

Sub Test_MinMax()
    Debug.Print ">> Test_MinMax"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("V"), Array(30), Array(10), Array(20))
    Assert "Max", df.MaxVal("V") = 30
    Assert "Min", df.MinVal("V") = 10
End Sub

Sub Test_Describe()
    Debug.Print ">> Test_Describe"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("N", "V"), Array("a", 10), Array("b", 20), Array("c", 30))
    Dim desc As DataFrame: Set desc = df.Describe
    Assert "Describe.Rows", desc.RowCount = 5
    Assert "Describe.HasV", desc.Value(1, "V") = 3  ' Count = 3
End Sub

Sub Test_GroupBy()
    Debug.Print ">> Test_GroupBy"
    Dim df As DataFrame
    Set df = DFrame.Create(Array("G", "V"), _
        Array("a", 10), Array("b", 20), Array("a", 30), Array("b", 40))
    
    Dim g As DataFrame: Set g = df.GroupBy("G").Sum("V")
    Assert "GroupBy.Rows", g.RowCount = 2
    ' Find "a" group
    Dim r As Long
    For r = 1 To g.RowCount
        If g.Value(r, "G") = "a" Then
            Assert "GroupBy.Sum_a", g.Value(r, "V_Sum") = 40
        End If
    Next r
    
    Dim gc As DataFrame: Set gc = df.GroupBy("G").Count
    Assert "GroupBy.Count.Cols", gc.ColCount = 2
End Sub

Sub Test_Join()
    Debug.Print ">> Test_Join"
    Dim left_ As DataFrame
    Set left_ = DFrame.Create(Array("K", "LV"), Array(1, "a"), Array(2, "b"), Array(3, "c"))
    Dim right_ As DataFrame
    Set right_ = DFrame.Create(Array("K", "RV"), Array(1, "x"), Array(3, "z"))
    
    Dim inner_ As DataFrame: Set inner_ = left_.JoinDF(right_, "K", "inner")
    Assert "Join.Inner.Rows", inner_.RowCount = 2
    
    Dim left2 As DataFrame: Set left2 = left_.JoinDF(right_, "K", "left")
    Assert "Join.Left.Rows", left2.RowCount = 3
End Sub

Sub Test_VStack()
    Debug.Print ">> Test_VStack"
    Dim df1 As DataFrame
    Set df1 = DFrame.Create(Array("A", "B"), Array(1, 2))
    Dim df2 As DataFrame
    Set df2 = DFrame.Create(Array("A", "B"), Array(3, 4), Array(5, 6))
    Dim merged As DataFrame: Set merged = df1.VStack(df2)
    Assert "VStack.Rows", merged.RowCount = 3
    Assert "VStack.Last", merged.Value(3, "A") = 5
End Sub

Sub Test_EmptyDF()
    Debug.Print ">> Test_EmptyDF"
    Dim df As DataFrame: Set df = DFrame.EmptyFrame(Array("A", "B"))
    Assert "Empty.Rows", df.RowCount = 0
    Assert "Empty.Cols", df.ColCount = 2
    
    ' Operations on empty DF should not error
    Assert "Empty.Head", df.Head(5).RowCount = 0
    Assert "Empty.Where", df.Where("A", "=", 1).RowCount = 0
End Sub
