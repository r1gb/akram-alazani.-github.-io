' Module: modDashboard
Option Compare Database
Option Explicit

' Fill dashboard controls (labels and subforms) with aggregated data
Public Sub LoadDashboard(frm As Form)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = CurrentDb

    ' Daily Sales (today)
    sql = "SELECT Nz(Sum([TotalAmount]),0) AS TodayTotal FROM tblSales WHERE DateValue([SaleDate])=Date()"
    Set rs = db.OpenRecordset(sql)
    If Not rs.EOF Then
        frm!lblTodaySales = rs!TodayTotal
    Else
        frm!lblTodaySales = 0
    End If
    rs.Close

    ' Monthly Sales (this month)
    sql = "SELECT Nz(Sum([TotalAmount]),0) AS MonthTotal FROM tblSales WHERE Year([SaleDate])=Year(Date()) AND Month([SaleDate])=Month(Date())"
    Set rs = db.OpenRecordset(sql)
    frm!lblMonthSales = IIf(rs.EOF,0,rs!MonthTotal)
    rs.Close

    ' Total Profit (last 30 days) - adjust field names if needed
    sql = "SELECT Nz(Sum(([UnitPrice]-[UnitCost])*[Quantity]),0) AS TotalProfit FROM tblInvoiceLines WHERE [InvoiceDate] >= DateAdd('d', -30, Date())"
    Set rs = db.OpenRecordset(sql)
    frm!lblProfit = IIf(rs.EOF,0,rs!TotalProfit)
    rs.Close

    ' Active customers
    sql = "SELECT Count(*) AS Cnt FROM tblCustomers WHERE Nz(IsActive,False)=True"
    Set rs = db.OpenRecordset(sql)
    frm!lblCustomers = IIf(rs.EOF,0,rs!Cnt)
    rs.Close

    ' Low stock count
    sql = "SELECT Count(*) AS LowCnt FROM tblProducts WHERE Nz(QtyOnHand,0) <= Nz(ReorderLevel,0)"
    Set rs = db.OpenRecordset(sql)
    frm!lblLowStock = IIf(rs.EOF,0,rs!LowCnt)
    rs.Close

    ' Optionally set subform RecordSource for latest sales
    frm!subLatestSales.Form.RecordSource = "SELECT TOP 10 SaleID, SaleDate, TotalAmount, CustomerID FROM tblSales ORDER BY SaleDate DESC"

ExitHandler:
    On Error Resume Next
    Set rs = Nothing: Set db = Nothing
    Exit Sub
ErrHandler:
    Debug.Print "LoadDashboard error: " & Err.Number & " - " & Err.Description
    Resume ExitHandler
End Sub

' Call this on Dashboard Form Load event:
' Private Sub Form_Load()
'     LoadDashboard Me
' End Sub
