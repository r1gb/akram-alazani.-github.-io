-- Sample SQL to create minimal tables and import sample data in MS Access
-- Run each CREATE TABLE in the SQL view of a new query and press Run.

CREATE TABLE tblProducts (
  ProductID COUNTER PRIMARY KEY,
  ProductName TEXT(255),
  QtyOnHand LONG,
  ReorderLevel LONG,
  UnitPrice CURRENCY,
  UnitCost CURRENCY
);

CREATE TABLE tblCustomers (
  CustomerID COUNTER PRIMARY KEY,
  CustomerName TEXT(255),
  IsActive YESNO,
  Email TEXT(255)
);

CREATE TABLE tblInvoices (
  InvoiceID COUNTER PRIMARY KEY,
  InvoiceNumber TEXT(50),
  SaleDate DATETIME,
  CustomerID LONG,
  TotalAmount CURRENCY
);

CREATE TABLE tblInvoiceLines (
  InvoiceLineID COUNTER PRIMARY KEY,
  InvoiceID LONG,
  ProductID LONG,
  Quantity LONG,
  UnitPrice CURRENCY,
  UnitCost CURRENCY
);

-- After creating tables, you can import the CSV files (External Data -> Text File) or use INSERT statements.
-- Example INSERT (for a few rows) -- adapt values as needed

INSERT INTO tblProducts (ProductName, QtyOnHand, ReorderLevel, UnitPrice, UnitCost) VALUES ("Keyboard",25,5,15.00,8.00);
INSERT INTO tblProducts (ProductName, QtyOnHand, ReorderLevel, UnitPrice, UnitCost) VALUES ("Mouse",40,10,8.50,3.50);

-- And similarly for Customers and Invoices.

-- To compute TotalAmount for existing invoices based on invoice lines, run:
UPDATE tblInvoices INNER JOIN (
    SELECT InvoiceID, Sum([Quantity]*[UnitPrice]) AS CalcTotal
    FROM tblInvoiceLines
    GROUP BY InvoiceID
) AS t ON tblInvoices.InvoiceID = t.InvoiceID
SET tblInvoices.TotalAmount = t.CalcTotal;
