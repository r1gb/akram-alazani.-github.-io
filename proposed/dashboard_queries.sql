-- Dashboard Queries (MS Access SQL)

-- 1) Daily Sales (last 30 days) -- adjust table/field names to your schema
SELECT DateValue([SaleDate]) AS SaleDate, SUM([TotalAmount]) AS DailyTotal
FROM tblSales
WHERE [SaleDate] BETWEEN Date() - 30 AND Date()
GROUP BY DateValue([SaleDate])
ORDER BY DateValue([SaleDate]) DESC;

-- 2) Monthly Sales (last 12 months)
SELECT Year([SaleDate]) AS SaleYear, Month([SaleDate]) AS SaleMonth,
       SUM([TotalAmount]) AS MonthlyTotal
FROM tblSales
WHERE [SaleDate] >= DateAdd('m', -12, Date())
GROUP BY Year([SaleDate]), Month([SaleDate])
ORDER BY Year([SaleDate]) DESC, Month([SaleDate]) DESC;

-- 3) Total Profit (last month) -- assume tblInvoiceLines with InvoiceDate, Quantity, UnitPrice, UnitCost
SELECT SUM(([UnitPrice] - [UnitCost]) * [Quantity]) AS TotalProfit
FROM tblInvoiceLines
WHERE [InvoiceDate] BETWEEN DateAdd('m', -1, Date()) AND Date();

-- 4) Active Customers
SELECT COUNT(*) AS ActiveCustomers FROM tblCustomers WHERE IsActive=True;

-- 5) Low stock products
SELECT p.ProductID, p.ProductName, p.QtyOnHand, p.ReorderLevel
FROM tblProducts AS p
WHERE p.QtyOnHand <= p.ReorderLevel
ORDER BY (p.QtyOnHand - p.ReorderLevel) ASC;
