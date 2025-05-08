Project Description:
Designed and implemented an interactive invoice template in Excel tailored for a landscaping service company. This project automates the invoicing process, reducing manual errors and improving efficiency. The template includes:
Auto-generated Invoice Number, PO Number, and Date fields. Input sections for customer details, itemized billing (description, quantity, price, VAT), and automatic subtotal, VAT, and total calculations. Integrated VBA buttons for:
Saving the invoice as Excel or PDF. Logging customer data for record-keeping. Resetting the form for the next customer. Space for customer messages, company info, and payment details for professional communication. Enhanced user interface with branding (logo) and easy navigation.
Outcome: Streamlined billing process for service businesses, allowing non-technical users to generate professional invoices with minimal input.

The Macro Code used is mentioned below

Option Explicit

Public InvoiceNumber As Long
Public CustomerName As String
Public Amount As Currency
Public DateIssued As Date
Public Terms As Byte

Sub RecordOfInvoice()

    Dim NextRecord As Range
    Dim a, b As Range

    InvoiceNumber = Range("C3")
    CustomerName = Range("B9")
    Amount = Range("I36")
    DateIssued = Range("C5")
    Terms = Range("C6")

    Set NextRecord = ActiveWorkbook.Worksheets("Record of Invoice").Range("A1048576").End(xlUp).Offset(1)
    Set a = ActiveWorkbook.Worksheets("Record of Invoice").Range("A1048576").End(xlUp)
    Set b = ActiveWorkbook.Worksheets("Invoice Template").Range("C3")

    If a <> b Then
        NextRecord = InvoiceNumber
        NextRecord.Offset(0, 1) = CustomerName
        NextRecord.Offset(0, 2) = Amount
        NextRecord.Offset(0, 3) = DateIssued
        NextRecord.Offset(0, 4) = DateIssued + Terms
    End If

    ActiveWorkbook.Save

End Sub

Sub SaveAsExcel()

    Dim FileLocation, File As String

    InvoiceNumber = Range("C3")
    CustomerName = Range("B9")

    Worksheets("Invoice Template").Copy

    Dim shp As Shape

    For Each shp In ActiveSheet.Shapes
        shp.Delete
    Next shp

    FileLocation = "E:\Data Analytics Projects"
    File = InvoiceNumber & "_" & CustomerName

    With ActiveWorkbook
        .Sheets(1).Name = "Invoice"
        .SaveAs Filename:=FileLocation & File, FileFormat:=51
        .Close
    End With

End Sub


Sub SaveAsPDF()

    Dim FileLocation, File As String

    InvoiceNumber = Range("C3")
    CustomerName = Range("B9")

    FileLocation = "E:\Data Analytics Projects\Excel Macro Project"
    File = InvoiceNumber & "_" & CustomerName

    ActiveSheet.Range("A1:I43").ExportAsFixedFormat Type:=xlTypePDF, Filename:=FileLocation & File & ".pdf"

End Sub

Sub NewInvoice()

    InvoiceNumber = Range("C3")
    Range("C3") = InvoiceNumber + 1
    Range("C4, C6, B9, B17:G31").ClearContents

    ActiveWorkbook.Save

End Sub

Sub NextCustomer()

    RecordOfInvoice
    SaveAsExcel
    SaveAsPDF
    NewInvoice

End Sub
