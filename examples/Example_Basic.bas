Attribute VB_Name = "Example_Basic"
Option Explicit

' ============================================================
'  VBA DataFrame - Basic Usage Examples
' ============================================================
'  Run each Sub individually from VBA Editor (F5)
'  Output appears in Immediate Window (Ctrl+G)
' ============================================================

' --- Example 1: Create from inline data ---
Sub Example_Create()
    Dim df As DataFrame
    Set df = DFrame.Create( _
        Array("Name", "Age", "City", "Sales"), _
        Array("Tanaka", 30, "Tokyo", 15000), _
        Array("Suzuki", 25, "Osaka", 8000), _
        Array("Sato", 35, "Tokyo", 22000), _
        Array("Yamada", 28, "Nagoya", 12000), _
        Array("Ito", 42, "Osaka", 31000), _
        Array("Takahashi", 33, "Tokyo", 18000))
    
    Debug.Print "=== DataFrame Created ==="
    Debug.Print df.Shape
    df.Print
End Sub

' --- Example 2: Read from Excel Range ---
Sub Example_FromRange()
    ' Assumes Sheet1 has data with headers in A1:D10
    Dim df As DataFrame
    Set df = DFrame.FromRange(Sheet1.Range("A1").CurrentRegion)
    
    Debug.Print "=== From Range ==="
    df.Head(5).Print
End Sub

' --- Example 3: Filtering ---
Sub Example_Filter()
    Dim df As DataFrame
    Set df = DFrame.Create( _
        Array("Product", "Category", "Price", "Stock"), _
        Array("Apple", "Fruit", 150, 200), _
        Array("Banana", "Fruit", 100, 350), _
        Array("Carrot", "Vegetable", 80, 150), _
        Array("Donut", "Sweets", 120, 50), _
        Array("Eggplant", "Vegetable", 200, 100))
    
    Debug.Print "=== All Data ==="
    df.Print
    
    Debug.Print vbCrLf & "=== Price > 100 ==="
    df.Where("Price", ">", 100).Print
    
    Debug.Print vbCrLf & "=== Category = Fruit ==="
    df.Where("Category", "=", "Fruit").Print
    
    Debug.Print vbCrLf & "=== Category In (Fruit, Sweets) ==="
    df.Where("Category", "In", Array("Fruit", "Sweets")).Print
End Sub

' --- Example 4: Method chaining ---
Sub Example_Chaining()
    Dim df As DataFrame
    Set df = DFrame.Create( _
        Array("Name", "Dept", "Salary"), _
        Array("A", "Sales", 5000), _
        Array("B", "Dev", 7000), _
        Array("C", "Sales", 6000), _
        Array("D", "Dev", 8000), _
        Array("E", "Sales", 4500), _
        Array("F", "HR", 5500))
    
    Debug.Print "=== Top 3 Salaries in Sales Dept ==="
    df.Where("Dept", "=", "Sales") _
      .OrderBy("Salary", Ascending:=False) _
      .Head(3) _
      .Print
End Sub

' --- Example 5: GroupBy ---
Sub Example_GroupBy()
    Dim df As DataFrame
    Set df = DFrame.Create( _
        Array("Region", "Product", "Sales", "Quantity"), _
        Array("East", "A", 1000, 10), _
        Array("West", "A", 1500, 15), _
        Array("East", "B", 2000, 20), _
        Array("West", "B", 800, 8), _
        Array("East", "A", 1200, 12), _
        Array("West", "B", 900, 9))
    
    Debug.Print "=== Sum by Region ==="
    df.GroupBy("Region").Sum("Sales", "Quantity").Print
    
    Debug.Print vbCrLf & "=== Mean by Region x Product ==="
    df.GroupBy("Region", "Product").Mean("Sales").Print
    
    Debug.Print vbCrLf & "=== Count by Region ==="
    df.GroupBy("Region").Count.Print
End Sub

' --- Example 6: Join ---
Sub Example_Join()
    Dim orders As DataFrame
    Set orders = DFrame.Create( _
        Array("OrderID", "CustID", "Amount"), _
        Array(1, "C01", 5000), _
        Array(2, "C02", 3000), _
        Array(3, "C01", 7000), _
        Array(4, "C03", 2000))
    
    Dim customers As DataFrame
    Set customers = DFrame.Create( _
        Array("CustID", "Name", "City"), _
        Array("C01", "Tanaka", "Tokyo"), _
        Array("C02", "Suzuki", "Osaka"), _
        Array("C04", "Yamada", "Nagoya"))
    
    Debug.Print "=== Inner Join ==="
    orders.JoinDF(customers, "CustID", "inner").Print
    
    Debug.Print vbCrLf & "=== Left Join ==="
    orders.JoinDF(customers, "CustID", "left").Print
End Sub

' --- Example 7: Describe ---
Sub Example_Describe()
    Dim df As DataFrame
    Set df = DFrame.Create( _
        Array("Name", "Score1", "Score2"), _
        Array("A", 80, 90), _
        Array("B", 65, 78), _
        Array("C", 92, 85), _
        Array("D", 71, 88), _
        Array("E", 55, 95))
    
    Debug.Print "=== Describe ==="
    df.Describe.Print
End Sub

' --- Example 8: Column operations ---
Sub Example_ColumnOps()
    Dim df As DataFrame
    Set df = DFrame.Create( _
        Array("Product", "Price", "Tax"), _
        Array("A", 1000, 100), _
        Array("B", 2000, 200), _
        Array("C", 1500, 150))
    
    Debug.Print "=== Original ==="
    df.Print
    
    ' Add column with scalar value
    Debug.Print vbCrLf & "=== Add Category column ==="
    df.AddCol("Category", "General").Print
    
    ' Remove column
    Debug.Print vbCrLf & "=== Remove Tax column ==="
    df.RemoveCol("Tax").Print
    
    ' Rename column
    Debug.Print vbCrLf & "=== Rename Price -> UnitPrice ==="
    df.RenameCol("Price", "UnitPrice").Print
End Sub

' --- Example 9: Write to Range ---
Sub Example_ToRange()
    Dim df As DataFrame
    Set df = DFrame.Create( _
        Array("Name", "Score"), _
        Array("Tanaka", 85), _
        Array("Suzuki", 92))
    
    ' Write to Sheet2 starting at A1
    df.ToRange Sheet2.Range("A1")
    Debug.Print "Data written to Sheet2!"
End Sub
