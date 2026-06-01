-- SQL: Security, Logs, Backups, Stock Alerts, Barcodes, Ledgers
-- MS Access (Jet) compatible SQL

-- 1) tblUsers
CREATE TABLE tblUsers (
  UserID COUNTER PRIMARY KEY,
  Username TEXT(50),
  PasswordHash TEXT(128),
  PasswordSalt TEXT(64),
  FullName TEXT(100),
  Email TEXT(100),
  IsActive YESNO,
  CreatedAt DATETIME,
  LastLogin DATETIME
);
CREATE INDEX idx_tblUsers_Username ON tblUsers (Username);

-- 2) Roles & Permissions
CREATE TABLE tblRoles (
  RoleID COUNTER PRIMARY KEY,
  RoleName TEXT(50),
  Description TEXT(255)
);

CREATE TABLE tblRolePermissions (
  RolePermissionID COUNTER PRIMARY KEY,
  RoleID LONG,
  PermissionKey TEXT(100), -- e.g. "Invoices.Create"
  Allow YESNO
);

CREATE TABLE tblUserRoles (
  UserRoleID COUNTER PRIMARY KEY,
  UserID LONG,
  RoleID LONG
);
CREATE INDEX idx_UserRoles_UserID ON tblUserRoles (UserID);

-- 3) Login Log
CREATE TABLE tblLoginLog (
  LoginLogID COUNTER PRIMARY KEY,
  UserID LONG,
  LoginTime DATETIME,
  LogoutTime DATETIME,
  Success YESNO,
  IPAddress TEXT(50),
  ComputerName TEXT(100),
  Notes MEMO
);
CREATE INDEX idx_LoginLog_UserID ON tblLoginLog (UserID);

-- 4) Audit Log
CREATE TABLE tblAuditLog (
  AuditID COUNTER PRIMARY KEY,
  UserID LONG,
  ActionTime DATETIME,
  ObjectType TEXT(50),
  ObjectID TEXT(50),
  Action TEXT(50),
  Details MEMO
);
CREATE INDEX idx_AuditLog_ActionTime ON tblAuditLog (ActionTime);

-- 5) Backups registry
CREATE TABLE tblBackups (
  BackupID COUNTER PRIMARY KEY,
  BackupPath TEXT(255),
  CreatedAt DATETIME,
  CreatedBy LONG,
  Notes MEMO
);

-- 6) Stock Alerts
CREATE TABLE tblStockAlerts (
  StockAlertID COUNTER PRIMARY KEY,
  ProductID LONG,
  Threshold LONG,
  Active YESNO,
  LastNotified DATETIME
);

-- 7) Product Barcodes
CREATE TABLE tblProductBarcodes (
  BarcodeID COUNTER PRIMARY KEY,
  ProductID LONG,
  BarcodeValue TEXT(50),
  BarcodeType TEXT(20)
);
CREATE INDEX idx_Barcode_Value ON tblProductBarcodes (BarcodeValue);

-- 8) Customer Ledger
CREATE TABLE tblCustomerLedger (
  LedgerID COUNTER PRIMARY KEY,
  CustomerID LONG,
  TransDate DATETIME,
  Reference TEXT(100),
  Debit CURRENCY,
  Credit CURRENCY,
  Balance CURRENCY,
  Notes MEMO
);
CREATE INDEX idx_CustLedger_CustomerID ON tblCustomerLedger (CustomerID);

-- 9) Supplier Ledger
CREATE TABLE tblSupplierLedger (
  LedgerID COUNTER PRIMARY KEY,
  SupplierID LONG,
  TransDate DATETIME,
  Reference TEXT(100),
  Debit CURRENCY,
  Credit CURRENCY,
  Balance CURRENCY,
  Notes MEMO
);
