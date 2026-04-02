VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "DataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  DataFrame  -  Pandas-like tabular data structure for VBA
' ============================================================

' --- Error constants ---
Private Const ERR_COL_NOT_FOUND As Long = &H80040401  'vbObjectError + 1025
Private Const ERR_COL_RANGE     As Long = &H80040402
Private Const ERR_DUP_COL       As Long = &H80040403
Private Const ERR_ROW_RANGE     As Long = &H80040404
Private Const ERR_INVALID_OP    As Long = &H80040405
Private Const ERR_COL_MISMATCH  As Long = &H80040406
Private Const ERR_EMPTY_DF      As Long = &H80040407
Private Const SRC               As String = "DataFrame"

' --- Internal storage ---
Private pData    As Variant       ' 2-D Variant array (1-based) or Empty
Private pCols()  As String        ' Column names (1-based)
Private pCI      As Object        ' Scripting.Dictionary  colName -> index
Private pRC      As Long          ' Row count
Private pCC      As Long          ' Column count

' ============================================================
'  Class events
' ============================================================
Private Sub Class_Initialize()
    Set pCI = CreateObject("Scripting.Dictionary")
    pCI.CompareMode = vbTextCompare
End Sub

Private Sub Class_Terminate()
    Set pCI = Nothing
End Sub

' ============================================================
'  Init  (called by factory functions in DF module)
' ============================================================
'  data     : 2-D Variant array (any base) or Empty for 0 rows
'  colNames : 1-D array of column name strings (any base)
Public Sub Init(ByRef data As Variant, ByRef colNames As Variant)
    Dim i As Long, ci As Long, r As Long, c As Long, dc As Long

    ' --- columns ---
    pCC = UBound(colNames) - LBound(colNames) + 1
    ReDim pCols(1 To pCC)
    pCI.RemoveAll
    For i = LBound(colNames) To UBound(colNames)
        ci = i - LBound(colNames) + 1
        pCols(ci) = CStr(colNames(i))
        If pCI.Exists(pCols(ci)) Then Err.Raise ERR_DUP_COL, SRC, "Duplicate column: " & pCols(ci)
        pCI.Add pCols(ci), ci
    Next i

    ' --- data ---
    If IsEmpty(data) Then
        pRC = 0: pData = Empty
        Exit Sub
    End If
    If Not IsArray(data) Then
        pRC = 0: pData = Empty
        Exit Sub
    End If

    pRC = UBound(data, 1) - LBound(data, 1) + 1
    dc = UBound(data, 2) - LBound(data, 2) + 1
    If dc <> pCC Then Err.Raise ERR_COL_MISMATCH, SRC, _
        "Data has " & dc & " cols but " & pCC & " names given"

    If LBound(data, 1) = 1 And LBound(data, 2) = 1 Then
        pData = data
    Else
        ReDim pData(1 To pRC, 1 To pCC)
        For r = 1 To pRC
            For c = 1 To pCC
                pData(r, c) = data(LBound(data, 1) + r - 1, LBound(data, 2) + c - 1)
            Next c
        Next r
    End If
End Sub

' ============================================================
'  Properties
' ============================================================
Public Property Get RowCount() As Long:  RowCount = pRC:  End Property
Public Property Get ColCount() As Long:  ColCount = pCC:  End Property

Public Property Get Shape() As String
    Shape = pRC & " rows x " & pCC & " cols"
End Property

Public Property Get Columns() As Variant
    Dim a() As String: ReDim a(1 To pCC)
    Dim i As Long
    For i = 1 To pCC: a(i) = pCols(i): Next i
    Columns = a
End Property

Public Property Get Values() As Variant: Values = pData: End Property

Public Property Get Value(ByVal rowIdx As Long, ByVal col As Variant) As Variant
    Dim c As Long: c = RCol(col)
    If rowIdx < 1 Or rowIdx > pRC Then Err.Raise ERR_ROW_RANGE, SRC, "Row " & rowIdx & " out of range"
    Value = pData(rowIdx, c)
End Property

' Return a single column as 1-D array
Public Function Col(ByVal colName As String) As Variant
    Dim c As Long, r As Long
    c = RCol(colName)
    If pRC = 0 Then Col = Array(): Exit Function
    Dim a() As Variant: ReDim a(1 To pRC)
    For r = 1 To pRC: a(r) = pData(r, c): Next r
    Col = a
End Function

' Return a single row as 1-D array
Public Function Row(ByVal rowIdx As Long) As Variant
    If rowIdx < 1 Or rowIdx > pRC Then Err.Raise ERR_ROW_RANGE, SRC, "Row " & rowIdx
    Dim a() As Variant: ReDim a(1 To pCC)
    Dim c As Long
    For c = 1 To pCC: a(c) = pData(rowIdx, c): Next c
    Row = a
End Function

' ============================================================
'  Private helpers
' ============================================================
Private Function RCol(ByVal col As Variant) As Long
    If VarType(col) = vbString Then
        If Not pCI.Exists(CStr(col)) Then Err.Raise ERR_COL_NOT_FOUND, SRC, "Column not found: " & col
        RCol = pCI(CStr(col))
    Else
        Dim idx As Long: idx = CLng(col)
        If idx < 1 Or idx > pCC Then Err.Raise ERR_COL_RANGE, SRC, "Column index out of range: " & col
        RCol = idx
    End If
End Function

Private Function CmpV(ByRef v1 As Variant, ByRef v2 As Variant) As Long
    If IsEmpty(v1) Or IsNull(v1) Then
        If IsEmpty(v2) Or IsNull(v2) Then CmpV = 0 Else CmpV = 1
        Exit Function
    End If
    If IsEmpty(v2) Or IsNull(v2) Then CmpV = -1: Exit Function
    If IsNumeric(v1) And IsNumeric(v2) Then
        Dim d1 As Double, d2 As Double: d1 = CDbl(v1): d2 = CDbl(v2)
        If d1 < d2 Then CmpV = -1 ElseIf d1 > d2 Then CmpV = 1 Else CmpV = 0
    Else
        CmpV = StrComp(CStr(v1), CStr(v2), vbTextCompare)
    End If
End Function

Private Sub QSort(ByRef ix() As Long, ByRef k As Variant, ByVal lo As Long, ByVal hi As Long, ByVal asc_ As Boolean)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long, tmp As Long, piv As Variant
    piv = k(ix((lo + hi) \ 2)): i = lo: j = hi
    Do While i <= j
        If asc_ Then
            Do While CmpV(k(ix(i)), piv) < 0: i = i + 1: Loop
            Do While CmpV(k(ix(j)), piv) > 0: j = j - 1: Loop
        Else
            Do While CmpV(k(ix(i)), piv) > 0: i = i + 1: Loop
            Do While CmpV(k(ix(j)), piv) < 0: j = j - 1: Loop
        End If
        If i <= j Then
            tmp = ix(i): ix(i) = ix(j): ix(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QSort ix, k, lo, j, asc_
    If i < hi Then QSort ix, k, i, hi, asc_
End Sub

' Build new DataFrame keeping all columns, selecting specific rows
Private Function RowSubset(ByRef ri() As Long, ByVal cnt As Long) As DataFrame
    Dim df As New DataFrame
    If cnt = 0 Then
        df.Init Empty, pCols: Set RowSubset = df: Exit Function
    End If
    Dim nd() As Variant: ReDim nd(1 To cnt, 1 To pCC)
    Dim r As Long, c As Long
    For r = 1 To cnt
        For c = 1 To pCC: nd(r, c) = pData(ri(r), c): Next c
    Next r
    df.Init nd, pCols
    Set RowSubset = df
End Function

' ============================================================
'  Selection methods
' ============================================================

' First n rows
Public Function Head(Optional ByVal n As Long = 5) As DataFrame
    If n > pRC Then n = pRC
    If n <= 0 Then
        Dim df As New DataFrame: df.Init Empty, pCols
        Set Head = df: Exit Function
    End If
    Dim ri() As Long: ReDim ri(1 To n)
    Dim i As Long
    For i = 1 To n: ri(i) = i: Next i
    Set Head = RowSubset(ri, n)
End Function

' Last n rows
Public Function Tail(Optional ByVal n As Long = 5) As DataFrame
    If n > pRC Then n = pRC
    If n <= 0 Then
        Dim df As New DataFrame: df.Init Empty, pCols
        Set Tail = df: Exit Function
    End If
    Dim ri() As Long: ReDim ri(1 To n)
    Dim i As Long, st As Long: st = pRC - n
    For i = 1 To n: ri(i) = st + i: Next i
    Set Tail = RowSubset(ri, n)
End Function

' Select columns by name  (use Sel because Select is reserved)
Public Function Sel(ParamArray cols() As Variant) As DataFrame
    Dim nC As Long: nC = UBound(cols) - LBound(cols) + 1
    Dim ci() As Long: ReDim ci(1 To nC)
    Dim cn() As String: ReDim cn(1 To nC)
    Dim i As Long
    For i = 0 To UBound(cols)
        ci(i + 1) = RCol(CStr(cols(i)))
        cn(i + 1) = pCols(ci(i + 1))
    Next i

    Dim df As New DataFrame
    If pRC = 0 Then
        df.Init Empty, cn: Set Sel = df: Exit Function
    End If
    Dim nd() As Variant: ReDim nd(1 To pRC, 1 To nC)
    Dim r As Long, c As Long
    For r = 1 To pRC
        For c = 1 To nC: nd(r, c) = pData(r, ci(c)): Next c
    Next r
    df.Init nd, cn
    Set Sel = df
End Function

' Row slice (1-based, inclusive)
Public Function Slice(ByVal startRow As Long, ByVal endRow As Long) As DataFrame
    If startRow < 1 Then startRow = 1
    If endRow > pRC Then endRow = pRC
    Dim cnt As Long: cnt = endRow - startRow + 1
    If cnt <= 0 Then
        Dim df As New DataFrame: df.Init Empty, pCols
        Set Slice = df: Exit Function
    End If
    Dim ri() As Long: ReDim ri(1 To cnt)
    Dim i As Long
    For i = 1 To cnt: ri(i) = startRow + i - 1: Next i
    Set Slice = RowSubset(ri, cnt)
End Function

' Filter rows:  op = "=","<>",">",">=","<","<=","Like","In"
Public Function Where(ByVal colName As String, ByVal op As String, ByVal Value As Variant) As DataFrame
    Dim ci As Long: ci = RCol(colName)
    Dim opU As String: opU = UCase$(Trim$(op))

    ' First pass: find matching row indices
    Dim ri() As Long: ReDim ri(1 To IIf(pRC > 0, pRC, 1))
    Dim cnt As Long: cnt = 0
    Dim r As Long, v As Variant, match As Boolean

    For r = 1 To pRC
        v = pData(r, ci): match = False
        Select Case opU
            Case "=":  match = (CmpV(v, Value) = 0)
            Case "<>": match = (CmpV(v, Value) <> 0)
            Case ">":  match = (CmpV(v, Value) > 0) 'Removed extra logic
            Case ">=": match = (CmpV(v, Value) >= 0)
            Case "<":  match = (CmpV(v, Value) < 0)
            Case "<=": match = (CmpV(v, Value) <= 0)
            Case "LIKE"
                If Not (IsEmpty(v) Or IsNull(v)) Then match = (CStr(v) Like CStr(Value))
            Case "IN"
                If IsArray(Value) Then
                    Dim vi As Long
                    For vi = LBound(Value) To UBound(Value)
                        If CmpV(v, Value(vi)) = 0 Then match = True: Exit For
                    Next vi
                End If
            Case Else
                Err.Raise ERR_INVALID_OP, SRC, "Unknown operator: " & op
        End Select
        If match Then cnt = cnt + 1: ri(cnt) = r
    Next r

    If cnt = 0 Then
        Dim df As New DataFrame: df.Init Empty, pCols
        Set Where = df
    Else
        ReDim Preserve ri(1 To cnt)
        Set Where = RowSubset(ri, cnt)
    End If
End Function

' Remove duplicate rows (or unique values of a column)
Public Function Distinct(Optional ByVal colName As String = "") As DataFrame
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ri() As Long: ReDim ri(1 To IIf(pRC > 0, pRC, 1))
    Dim cnt As Long: cnt = 0
    Dim r As Long, c As Long, key As String

    If colName = "" Then
        ' Distinct on all columns
        For r = 1 To pRC
            key = ""
            For c = 1 To pCC: key = key & CStr(pData(r, c) & "") & Chr$(0): Next c
            If Not dict.Exists(key) Then
                dict.Add key, 1: cnt = cnt + 1: ri(cnt) = r
            End If
        Next r
    Else
        Dim ci As Long: ci = RCol(colName)
        For r = 1 To pRC
            key = CStr(pData(r, ci) & "")
            If Not dict.Exists(key) Then
                dict.Add key, 1: cnt = cnt + 1: ri(cnt) = r
            End If
        Next r
    End If

    If cnt = 0 Then
        Dim df As New DataFrame: df.Init Empty, pCols
        Set Distinct = df
    Else
        ReDim Preserve ri(1 To cnt)
        Set Distinct = RowSubset(ri, cnt)
    End If
End Function

' ============================================================
'  Transformation methods
' ============================================================

' Sort by column
Public Function OrderBy(ByVal colName As String, Optional ByVal Ascending As Boolean = True) As DataFrame
    If pRC <= 1 Then Set OrderBy = Me.Head(pRC): Exit Function
    Dim ci As Long: ci = RCol(colName)
    Dim keys() As Variant: ReDim keys(1 To pRC)
    Dim ix() As Long: ReDim ix(1 To pRC)
    Dim r As Long
    For r = 1 To pRC: keys(r) = pData(r, ci): ix(r) = r: Next r
    QSort ix, keys, 1, pRC, Ascending
    Set OrderBy = RowSubset(ix, pRC)
End Function

' Add a column (returns new DataFrame)
' vals: 1-D array of values (same length as RowCount) or a scalar
Public Function AddCol(ByVal colName As String, ByVal vals As Variant) As DataFrame
    Dim newCols() As String: ReDim newCols(1 To pCC + 1)
    Dim i As Long
    For i = 1 To pCC: newCols(i) = pCols(i): Next i
    newCols(pCC + 1) = colName

    Dim df As New DataFrame
    If pRC = 0 Then df.Init Empty, newCols: Set AddCol = df: Exit Function

    Dim nd() As Variant: ReDim nd(1 To pRC, 1 To pCC + 1)
    Dim r As Long, c As Long
    For r = 1 To pRC
        For c = 1 To pCC: nd(r, c) = pData(r, c): Next c
        If IsArray(vals) Then
            nd(r, pCC + 1) = vals(LBound(vals) + r - 1)
        Else
            nd(r, pCC + 1) = vals
        End If
    Next r
    df.Init nd, newCols
    Set AddCol = df
End Function

' Remove a column
Public Function RemoveCol(ByVal colName As String) As DataFrame
    Dim ci As Long: ci = RCol(colName)
    If pCC = 1 Then Err.Raise ERR_COL_MISMATCH, SRC, "Cannot remove the only column"

    Dim newCols() As String: ReDim newCols(1 To pCC - 1)
    Dim cMap() As Long: ReDim cMap(1 To pCC - 1)
    Dim i As Long, j As Long: j = 0
    For i = 1 To pCC
        If i <> ci Then j = j + 1: newCols(j) = pCols(i): cMap(j) = i
    Next i

    Dim df As New DataFrame
    If pRC = 0 Then df.Init Empty, newCols: Set RemoveCol = df: Exit Function

    Dim nd() As Variant: ReDim nd(1 To pRC, 1 To pCC - 1)
    Dim r As Long, c As Long
    For r = 1 To pRC
        For c = 1 To pCC - 1: nd(r, c) = pData(r, cMap(c)): Next c
    Next r
    df.Init nd, newCols
    Set RemoveCol = df
End Function

' Rename a column
Public Function RenameCol(ByVal oldName As String, ByVal newName As String) As DataFrame
    Dim ci As Long: ci = RCol(oldName)
    Dim newCols() As String: ReDim newCols(1 To pCC)
    Dim i As Long
    For i = 1 To pCC: newCols(i) = pCols(i): Next i
    newCols(ci) = newName

    Dim df As New DataFrame
    If pRC = 0 Then df.Init Empty, newCols: Set RenameCol = df: Exit Function
    df.Init pData, newCols
    Set RenameCol = df
End Function

' ============================================================
'  Aggregation helpers (single column)
' ============================================================
Public Function Sum(ByVal colName As String) As Double
    Dim ci As Long: ci = RCol(colName)
    Dim r As Long, s As Double
    For r = 1 To pRC
        If IsNumeric(pData(r, ci)) Then s = s + CDbl(pData(r, ci))
    Next r
    Sum = s
End Function

Public Function Mean(ByVal colName As String) As Double
    Dim ci As Long: ci = RCol(colName)
    Dim r As Long, s As Double, n As Long
    For r = 1 To pRC
        If IsNumeric(pData(r, ci)) Then s = s + CDbl(pData(r, ci)): n = n + 1
    Next r
    If n > 0 Then Mean = s / n
End Function

Public Function MaxVal(ByVal colName As String) As Variant
    Dim ci As Long: ci = RCol(colName)
    Dim r As Long, mx As Variant: mx = Empty
    For r = 1 To pRC
        If IsEmpty(mx) Then mx = pData(r, ci) _
        ElseIf CmpV(pData(r, ci), mx) < 0 Then ' current < mx => skip
        ElseIf CmpV(pData(r, ci), mx) > 0 Then mx = pData(r, ci)
        End If
    Next r
    MaxVal = mx
End Function

Public Function MinVal(ByVal colName As String) As Variant
    Dim ci As Long: ci = RCol(colName)
    Dim r As Long, mn As Variant: mn = Empty
    For r = 1 To pRC
        If IsEmpty(mn) Then mn = pData(r, ci) _
        ElseIf CmpV(pData(r, ci), mn) > 0 Then ' current > mn => skip
        ElseIf CmpV(pData(r, ci), mn) < 0 Then mn = pData(r, ci)
        End If
    Next r
    MinVal = mn
End Function

' Basic statistics table
Public Function Describe() As DataFrame
    ' Detect numeric columns
    Dim isNum() As Boolean: ReDim isNum(1 To pCC)
    Dim numCnt As Long, i As Long, r As Long
    For i = 1 To pCC
        Dim nCnt As Long: nCnt = 0
        For r = 1 To pRC
            If IsNumeric(pData(r, i)) And Not IsEmpty(pData(r, i)) Then nCnt = nCnt + 1
        Next r
        isNum(i) = (nCnt > 0)
        If isNum(i) Then numCnt = numCnt + 1
    Next i

    If numCnt = 0 Then
        Dim df0 As New DataFrame: df0.Init Empty, Array("Stat")
        Set Describe = df0: Exit Function
    End If

    ' Build column list
    Dim cn() As String: ReDim cn(1 To numCnt + 1)
    Dim cIdx() As Long: ReDim cIdx(1 To numCnt)
    cn(1) = "Stat": Dim j As Long: j = 0
    For i = 1 To pCC
        If isNum(i) Then j = j + 1: cn(j + 1) = pCols(i): cIdx(j) = i
    Next i

    ' 5 stats: Count, Sum, Mean, Min, Max
    Dim nd() As Variant: ReDim nd(1 To 5, 1 To numCnt + 1)
    nd(1, 1) = "Count": nd(2, 1) = "Sum": nd(3, 1) = "Mean"
    nd(4, 1) = "Min": nd(5, 1) = "Max"

    For j = 1 To numCnt
        Dim ci As Long: ci = cIdx(j)
        Dim s As Double, n As Long, mx As Variant, mn As Variant
        s = 0: n = 0: mx = Empty: mn = Empty
        For r = 1 To pRC
            If IsNumeric(pData(r, ci)) And Not IsEmpty(pData(r, ci)) Then
                Dim v As Double: v = CDbl(pData(r, ci))
                s = s + v: n = n + 1
                If IsEmpty(mx) Then mx = v: mn = v
                If v > CDbl(mx) Then mx = v
                If v < CDbl(mn) Then mn = v
            End If
        Next r
        nd(1, j + 1) = n
        nd(2, j + 1) = s
        nd(3, j + 1) = IIf(n > 0, s / n, Empty)
        nd(4, j + 1) = mn
        nd(5, j + 1) = mx
    Next j

    Dim dfR As New DataFrame: dfR.Init nd, cn
    Set Describe = dfR
End Function

' GroupBy — returns DataFrameGroupBy object
Public Function GroupBy(ParamArray cols() As Variant) As DataFrameGroupBy
    Dim grp As New DataFrameGroupBy
    Dim arr As Variant: arr = cols   ' capture ParamArray
    grp.Init Me, arr
    Set GroupBy = grp
End Function

' ============================================================
'  Join
' ============================================================

' Join two DataFrames on a common column
' How: "inner" (default), "left"
Public Function JoinDF(ByRef other As DataFrame, ByVal OnCol As String, _
                       Optional ByVal How As String = "inner") As DataFrame
    Dim lci As Long: lci = RCol(OnCol)
    Dim howU As String: howU = LCase$(How)

    ' Build right-side lookup:  key -> Collection of row indices
    Dim rLook As Object: Set rLook = CreateObject("Scripting.Dictionary")
    rLook.CompareMode = vbTextCompare
    Dim r As Long, key As String
    For r = 1 To other.RowCount
        key = CStr(other.Value(r, OnCol) & "")
        If Not rLook.Exists(key) Then rLook.Add key, New Collection
        rLook(key).Add r
    Next r

    ' Determine result columns: all left cols + right cols excluding OnCol
    Dim rCols As Variant: rCols = other.Columns
    Dim rKeep() As Long, rNames() As String
    Dim rCnt As Long: rCnt = 0
    Dim i As Long
    For i = 1 To other.ColCount
        If StrComp(CStr(rCols(i)), OnCol, vbTextCompare) <> 0 Then rCnt = rCnt + 1
    Next i
    ReDim rKeep(1 To IIf(rCnt > 0, rCnt, 1))
    ReDim rNames(1 To IIf(rCnt > 0, rCnt, 1))
    Dim k As Long: k = 0
    For i = 1 To other.ColCount
        If StrComp(CStr(rCols(i)), OnCol, vbTextCompare) <> 0 Then
            k = k + 1: rKeep(k) = i: rNames(k) = CStr(rCols(i))
        End If
    Next i

    Dim nC As Long: nC = pCC + rCnt
    Dim resCols() As String: ReDim resCols(1 To nC)
    For i = 1 To pCC: resCols(i) = pCols(i): Next i
    For i = 1 To rCnt: resCols(pCC + i) = rNames(i): Next i

    ' Collect result rows (two passes: count then fill)
    ' Pass 1: count
    Dim totalRows As Long: totalRows = 0
    For r = 1 To pRC
        key = CStr(pData(r, lci) & "")
        If rLook.Exists(key) Then
            totalRows = totalRows + rLook(key).Count
        ElseIf howU = "left" Then
            totalRows = totalRows + 1
        End If
    Next r

    Dim df As New DataFrame
    If totalRows = 0 Then
        df.Init Empty, resCols: Set JoinDF = df: Exit Function
    End If

    ' Pass 2: fill
    Dim nd() As Variant: ReDim nd(1 To totalRows, 1 To nC)
    Dim ri As Long: ri = 0
    Dim c As Long, rr As Long
    For r = 1 To pRC
        key = CStr(pData(r, lci) & "")
        If rLook.Exists(key) Then
            Dim coll As Collection: Set coll = rLook(key)
            Dim m As Long
            For m = 1 To coll.Count
                ri = ri + 1
                For c = 1 To pCC: nd(ri, c) = pData(r, c): Next c
                rr = coll(m)
                For c = 1 To rCnt: nd(ri, pCC + c) = other.Value(rr, CLng(rKeep(c))): Next c
            Next m
        ElseIf howU = "left" Then
            ri = ri + 1
            For c = 1 To pCC: nd(ri, c) = pData(r, c): Next c
            For c = 1 To rCnt: nd(ri, pCC + c) = Empty: Next c
        End If
    Next r

    df.Init nd, resCols
    Set JoinDF = df
End Function

' Vertical stack (union)
Public Function VStack(ByRef other As DataFrame) As DataFrame
    If pCC <> other.ColCount Then Err.Raise ERR_COL_MISMATCH, SRC, "Column count mismatch"
    Dim total As Long: total = pRC + other.RowCount
    Dim df As New DataFrame
    If total = 0 Then df.Init Empty, pCols: Set VStack = df: Exit Function

    Dim nd() As Variant: ReDim nd(1 To total, 1 To pCC)
    Dim r As Long, c As Long
    For r = 1 To pRC
        For c = 1 To pCC: nd(r, c) = pData(r, c): Next c
    Next r
    For r = 1 To other.RowCount
        For c = 1 To pCC: nd(pRC + r, c) = other.Value(r, CLng(c)): Next c
    Next r
    df.Init nd, pCols
    Set VStack = df
End Function

' ============================================================
'  Output
' ============================================================

' Write to Excel range (header + data)
Public Sub ToRange(ByVal target As Range)
    Dim c As Long
    For c = 1 To pCC: target.Cells(1, c).Value = pCols(c): Next c
    If pRC > 0 Then target.Cells(2, 1).Resize(pRC, pCC).Value = pData
End Sub

' Return data as 2-D array
Public Function ToArray() As Variant
    ToArray = pData
End Function

' Print to Immediate Window
Public Sub Print(Optional ByVal maxRows As Long = 20)
    Debug.Print ToString(maxRows)
End Sub

' Format as string table
Public Function ToString(Optional ByVal maxRows As Long = 0) As String
    If maxRows <= 0 Then maxRows = pRC

    Dim showRows As Long: showRows = pRC
    Dim truncated As Boolean: truncated = False
    If pRC > maxRows Then showRows = maxRows: truncated = True

    ' Calculate column widths
    Dim w() As Long: ReDim w(1 To pCC)
    Dim i As Long, r As Long, sLen As Long
    For i = 1 To pCC
        w(i) = Len(pCols(i))
        For r = 1 To showRows
            sLen = Len(CStr(pData(r, i) & ""))
            If sLen > w(i) Then w(i) = sLen
        Next r
        If w(i) > 30 Then w(i) = 30  ' cap width
    Next i

    ' Build output
    Dim lines() As String
    ReDim lines(0 To showRows + 1 + IIf(truncated, 1, 0))
    Dim lIdx As Long: lIdx = 0

    ' Header
    Dim hdr As String, sep As String
    For i = 1 To pCC
        hdr = hdr & PadR(pCols(i), w(i))
        sep = sep & String$(w(i), "-")
        If i < pCC Then hdr = hdr & "  ": sep = sep & "  "
    Next i
    lines(0) = hdr: lines(1) = sep: lIdx = 2

    ' Data rows
    Dim ln As String
    For r = 1 To showRows
        ln = ""
        For i = 1 To pCC
            ln = ln & PadR(CStr(pData(r, i) & ""), w(i))
            If i < pCC Then ln = ln & "  "
        Next i
        lines(lIdx) = ln: lIdx = lIdx + 1
    Next r

    If truncated Then lines(lIdx) = "... (" & pRC & " rows total)"

    ToString = Join(lines, vbCrLf)
End Function

Private Function PadR(ByVal s As String, ByVal w As Long) As String
    If Len(s) >= w Then PadR = Left$(s, w) Else PadR = s & Space$(w - Len(s))
End Function
