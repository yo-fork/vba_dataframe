Attribute VB_Name = "DFrame"
Option Explicit

' ============================================================
'  DFrame  -  Factory functions for DataFrame
' ============================================================
'  Usage:
'    Dim df As DataFrame
'    Set df = DFrame.FromRange(Sheet1.Range("A1:D100"))
'    Set df = DFrame.Create(Array("Name","Age"), Array("Taro",30), Array("Hanako",25))
' ============================================================

' Create DataFrame from an Excel Range
' If hasHeader=True (default), first row is used as column names
Public Function FromRange(ByVal rng As Range, _
                          Optional ByVal hasHeader As Boolean = True) As DataFrame
    Dim df As New DataFrame
    Dim data As Variant
    Dim colNames() As String
    Dim nCols As Long: nCols = rng.Columns.Count
    Dim nRows As Long: nRows = rng.Rows.Count
    Dim i As Long

    ReDim colNames(1 To nCols)

    If hasHeader Then
        ' Extract column names from first row
        For i = 1 To nCols
            colNames(i) = CStr(rng.Cells(1, i).Value2)
            If colNames(i) = "" Then colNames(i) = "Col" & i
        Next i

        If nRows <= 1 Then
            ' Header only, no data
            df.Init Empty, colNames
        Else
            data = rng.Offset(1, 0).Resize(nRows - 1, nCols).Value2
            ' Handle single-cell case (Value2 returns scalar, not array)
            If Not IsArray(data) Then
                Dim tmp(1 To 1, 1 To 1) As Variant
                tmp(1, 1) = data: data = tmp
            End If
            df.Init data, colNames
        End If
    Else
        ' Auto-generate column names
        For i = 1 To nCols: colNames(i) = "Col" & i: Next i

        If nRows = 1 And nCols = 1 Then
            Dim tmp2(1 To 1, 1 To 1) As Variant
            tmp2(1, 1) = rng.Value2: data = tmp2
        Else
            data = rng.Value2
        End If
        df.Init data, colNames
    End If

    Set FromRange = df
End Function

' Create DataFrame from a 2-D array
' colNames: 1-D array of column name strings
Public Function FromArray(ByRef data As Variant, ByRef colNames As Variant) As DataFrame
    Dim df As New DataFrame
    df.Init data, colNames
    Set FromArray = df
End Function

' Create DataFrame from inline data
' First argument: Array of column names
' Remaining arguments: each is an Array representing one row
'   Example: DFrame.Create(Array("A","B"), Array(1,2), Array(3,4))
Public Function Create(ByVal colNames As Variant, ParamArray rows() As Variant) As DataFrame
    Dim df As New DataFrame
    Dim nCols As Long: nCols = UBound(colNames) - LBound(colNames) + 1
    Dim nRows As Long: nRows = UBound(rows) - LBound(rows) + 1

    If nRows = 0 Then
        df.Init Empty, colNames
        Set Create = df
        Exit Function
    End If

    Dim data() As Variant: ReDim data(1 To nRows, 1 To nCols)
    Dim r As Long, c As Long, rowArr As Variant

    For r = LBound(rows) To UBound(rows)
        rowArr = rows(r)
        For c = LBound(rowArr) To UBound(rowArr)
            data(r - LBound(rows) + 1, c - LBound(rowArr) + 1) = rowArr(c)
        Next c
    Next r

    df.Init data, colNames
    Set Create = df
End Function

' Create an empty DataFrame with specified columns
Public Function EmptyFrame(ByVal colNames As Variant) As DataFrame
    Dim df As New DataFrame
    df.Init Empty, colNames
    Set EmptyFrame = df
End Function

' Create DataFrame from a CSV file
Public Function FromCsv(ByVal filePath As String, _
                        Optional ByVal sep As String = ",", _
                        Optional ByVal hasHeader As Boolean = True) As DataFrame
    Dim fNum As Integer: fNum = FreeFile
    Dim lines() As String
    Dim content As String
    Dim lineCount As Long, colCount As Long
    Dim i As Long, j As Long

    ' Read entire file
    Open filePath For Input As #fNum
    content = Input$(LOF(fNum), fNum)
    Close #fNum

    ' Strip BOM if present (UTF-8 BOM appears as Chr(&HFEFF) or 3-byte EF BB BF)
    If Len(content) > 0 Then
        If AscW(Left$(content, 1)) = &HFEFF Then
            content = Mid$(content, 2)
        ElseIf Len(content) >= 3 Then
            If Mid$(content, 1, 3) = Chr$(&HEF) & Chr$(&HBB) & Chr$(&HBF) Then
                content = Mid$(content, 4)
            End If
        End If
    End If

    ' Split into lines (handle both CR+LF and LF)
    content = Replace(content, vbCr, "")
    lines = Split(content, vbLf)

    ' Remove trailing empty lines
    lineCount = UBound(lines)
    Do While lineCount >= 0 And Trim$(lines(lineCount)) = ""
        lineCount = lineCount - 1
    Loop
    lineCount = lineCount + 1  ' total non-empty lines

    If lineCount = 0 Then
        Dim dfE As New DataFrame
        dfE.Init Empty, Array("Col1")
        Set FromCsv = dfE
        Exit Function
    End If

    ' Parse header
    Dim headerFields() As String
    headerFields = Split(lines(0), sep)
    colCount = UBound(headerFields) + 1

    Dim colNames() As String: ReDim colNames(1 To colCount)
    If hasHeader Then
        For i = 0 To UBound(headerFields)
            colNames(i + 1) = Trim$(headerFields(i))
            If colNames(i + 1) = "" Then colNames(i + 1) = "Col" & (i + 1)
        Next i
    Else
        For i = 1 To colCount: colNames(i) = "Col" & i: Next i
    End If

    ' Parse data rows
    Dim dataStart As Long
    If hasHeader Then dataStart = 1 Else dataStart = 0
    Dim nRows As Long: nRows = lineCount - dataStart

    If nRows <= 0 Then
        Dim df0 As New DataFrame
        df0.Init Empty, colNames
        Set FromCsv = df0
        Exit Function
    End If

    Dim data() As Variant: ReDim data(1 To nRows, 1 To colCount)
    Dim fields() As String
    Dim r As Long
    Dim fVal As String

    For i = dataStart To lineCount - 1
        r = i - dataStart + 1
        fields = Split(lines(i), sep)
        For j = 0 To UBound(fields)
            If j + 1 <= colCount Then
                fVal = Trim$(fields(j))
                ' Try numeric conversion
                If IsNumeric(fVal) And fVal <> "" Then
                    data(r, j + 1) = CDbl(fVal)
                Else
                    data(r, j + 1) = fVal
                End If
            End If
        Next j
    Next i

    Dim df As New DataFrame
    df.Init data, colNames
    Set FromCsv = df
End Function
