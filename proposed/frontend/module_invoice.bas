' Module: modInvoiceAndBarcode
Option Compare Database
Option Explicit

' Generate barcode string for Code128 (simple mapping) or use barcode font.
' For printing, it's often easiest to install a barcode font (e.g., Code128) and set control font.

Public Function FormatCode128(ByVal data As String) As String
    ' Placeholder: real Code128 encoding requires checksum and start/stop characters.
    ' If using a barcode font that requires start/stop (e.g., * for Code39), wrap accordingly.
    FormatCode128 = "*" & data & "*"
End Function

' Prepare invoice form/report: set RecordSource and fill controls
Public Sub PrepareInvoice(frm As Form, ByVal InvoiceID As Long)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb

    ' Set main recordsource to invoice header
    frm.RecordSource = "SELECT * FROM tblInvoices WHERE InvoiceID = " & InvoiceID

    ' Set subform (lines) to invoice lines
    frm!subInvoiceLines.Form.RecordSource = "SELECT InvoiceLineID, ProductID, Quantity, UnitPrice, (Quantity*UnitPrice) AS LineTotal FROM tblInvoiceLines WHERE InvoiceID = " & InvoiceID

    ' Load barcode value into control (assuming tblInvoices.InvoiceNumber exists)
    Set rs = db.OpenRecordset("SELECT InvoiceNumber FROM tblInvoices WHERE InvoiceID = " & InvoiceID)
    If Not rs.EOF Then
        frm!txtBarcode = FormatCode128(Nz(rs!InvoiceNumber, ""))
        ' Set font of txtBarcode to barcode font (user must install Code128 or similar)
        ' frm!txtBarcode.FontName = "IDAutomationHC39M" ' example
    End If
    rs.Close

ExitHandler:
    On Error Resume Next
    Set rs = Nothing: Set db = Nothing
    Exit Sub
ErrHandler:
    Debug.Print "PrepareInvoice error: " & Err.Number & " - " & Err.Description
    Resume ExitHandler
End Sub

' Printing example
Public Sub PrintInvoice(ByVal InvoiceID As Long)
    DoCmd.OpenForm "frmInvoice", acNormal
    DoCmd.RunCommand acCmdRefresh
    ' Pass InvoiceID to opened form
    Forms!frmInvoice!InvoiceID = InvoiceID
    PrepareInvoice Forms!frmInvoice, InvoiceID
    DoCmd.PrintOut
    DoCmd.Close acForm, "frmInvoice"
End Sub
