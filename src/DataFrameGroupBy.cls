VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "DataFrameGroupBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  DataFrameGroupBy  -  GroupBy aggregation helper
' ============================================================

Private pSource      As DataFrame
Private pGroupCols() As String
Private pGC          As Long           ' number of group columns
Private pGroups      As Object         ' Dict: groupKey -> Collection of row indices
Private pGroupKeys() As Variant        ' array of group key arrays (for ordered iteration)
Private pGroupCount  As Long

' Called from DataFrame.GroupBy
Public Sub Init(ByRef src As DataFrame, ByRef cols As Variant)
    Set pSource = src
    Dim i As Long

    ' Unpack column names
    pGC = UBound(cols) - LBound(cols) + 1
    ReDim pGroupCols(1 To pGC)
    For i = LBound(cols) To UBound(cols)
        pGroupCols(i - LBound(cols) + 1) = CStr(cols(i))
    Next i

    ' Build groups
    Set pGroups = CreateObject("Scripting.Dictionary")
    pGroups.CompareMode = vbTextCompare

    Dim r As Long, key As String, j As Long
    Dim keyVals() As Variant
    pGroupCount = 0

    For r = 1 To pSource.RowCount
        key = ""
        ReDim keyVals(1 To pGC)
        For j = 1 To pGC
            keyVals(j) = pSource.Value(r, pGroupCols(j))
            key = key & CStr(keyVals(j) & "") & Chr$(0)
        Next j
        If Not pGroups.Exists(key) Then
            pGroupCount = pGroupCount + 1
            Dim c As New Collection
            pGroups.Add key, Array(keyVals, c)
        End If
        Dim arr As Variant: arr = pGroups(key)
        arr(1).Add r
        pGroups(key) = arr   ' reassign since array is copied
    Next r
End Sub

Public Property Get GroupCount() As Long: GroupCount = pGroupCount: End Property

' ============================================================
'  Aggregation methods
' ============================================================

' Sum of specified columns per group
Public Function Sum(ParamArray cols() As Variant) As DataFrame
    Sum = Aggregate("Sum", cols)
End Function

' Mean of specified columns per group
Public Function Mean(ParamArray cols() As Variant) As DataFrame
    Mean = Aggregate("Mean", cols)
End Function

' Count per group
Public Function Count() As DataFrame
    Dim dummy As Variant: dummy = Array()
    Set Count = Aggregate("Count", dummy)
End Function

' Max of specified columns per group
Public Function MaxVal(ParamArray cols() As Variant) As DataFrame
    Set MaxVal = Aggregate("Max", cols)
End Function

' Min of specified columns per group
Public Function MinVal(ParamArray cols() As Variant) As DataFrame
    Set MinVal = Aggregate("Min", cols)
End Function

' ============================================================
'  Internal aggregation engine
' ============================================================
Private Function Aggregate(ByVal funcName As String, ByRef cols As Variant) As DataFrame
    Dim aggCols() As String
    Dim nAgg As Long
    Dim i As Long, j As Long

    ' Determine aggregation columns
    If funcName = "Count" Then
        nAgg = 0
    Else
        nAgg = UBound(cols) - LBound(cols) + 1
        ReDim aggCols(1 To nAgg)
        For i = LBound(cols) To UBound(cols)
            aggCols(i - LBound(cols) + 1) = CStr(cols(i))
        Next i
    End If

    ' Build result column names
    Dim nResultCols As Long
    If funcName = "Count" Then
        nResultCols = pGC + 1
    Else
        nResultCols = pGC + nAgg
    End If

    Dim resCols() As String: ReDim resCols(1 To nResultCols)
    For i = 1 To pGC: resCols(i) = pGroupCols(i): Next i
    If funcName = "Count" Then
        resCols(pGC + 1) = "Count"
    Else
        For i = 1 To nAgg: resCols(pGC + i) = aggCols(i) & "_" & funcName: Next i
    End If

    ' Build result data
    Dim nd() As Variant: ReDim nd(1 To pGroupCount, 1 To nResultCols)
    Dim keys As Variant: keys = pGroups.keys
    Dim gIdx As Long

    For gIdx = 0 To UBound(keys)
        Dim gKey As String: gKey = keys(gIdx)
        Dim gArr As Variant: gArr = pGroups(gKey)
        Dim gKeyVals As Variant: gKeyVals = gArr(0)
        Dim gRows As Collection: Set gRows = gArr(1)

        ' Fill group key columns
        For j = 1 To pGC
            nd(gIdx + 1, j) = gKeyVals(j)
        Next j

        ' Compute aggregates
        If funcName = "Count" Then
            nd(gIdx + 1, pGC + 1) = gRows.Count
        Else
            For i = 1 To nAgg
                Dim colResult As Variant
                colResult = ComputeAgg(funcName, aggCols(i), gRows)
                nd(gIdx + 1, pGC + i) = colResult
            Next i
        End If
    Next gIdx

    Dim df As New DataFrame
    df.Init nd, resCols
    Set Aggregate = df
End Function

Private Function ComputeAgg(ByVal funcName As String, ByVal colName As String, _
                            ByRef rows As Collection) As Variant
    Dim r As Long, v As Variant, val As Variant
    Dim s As Double, n As Long, mx As Variant, mn As Variant
    s = 0: n = 0: mx = Empty: mn = Empty

    For r = 1 To rows.Count
        v = pSource.Value(rows(r), colName)
        If IsNumeric(v) And Not IsEmpty(v) And Not IsNull(v) Then
            val = CDbl(v)
            s = s + val: n = n + 1
            If IsEmpty(mx) Then mx = val: mn = val
            If val > CDbl(mx) Then mx = val
            If val < CDbl(mn) Then mn = val
        End If
    Next r

    Select Case funcName
        Case "Sum":  ComputeAgg = s
        Case "Mean": ComputeAgg = IIf(n > 0, s / n, Empty)
        Case "Max":  ComputeAgg = mx
        Case "Min":  ComputeAgg = mn
        Case Else:   ComputeAgg = Empty
    End Select
End Function
