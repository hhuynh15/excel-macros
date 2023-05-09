Attribute VB_Name = "Module1"
Option Explicit

Sub SendEmails()
    Dim ws As Worksheet
    Dim OutApp As Object
    Dim OutMail As Object
    Dim i As Long
    Dim lastRow As Long
    Dim recipientEmail As String
    Dim recipientName As String
    Dim EmailsDict As Object
    Dim emailKey As Variant

    Set ws = ThisWorkbook.Sheets("Sheet1") 'Change "Sheet1" to the name of the sheet with the data
    Set OutApp = CreateObject("Outlook.Application")
    Set EmailsDict = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, "X").End(xlUp).Row

    ' Collect unique email addresses
    For i = 2 To lastRow
        recipientEmail = ws.Cells(i, "X").Value
        If recipientEmail <> "" And Not EmailsDict.Exists(recipientEmail) Then
            EmailsDict.Add recipientEmail, i
        End If
    Next i

    ' Process each unique email address
    For Each emailKey In EmailsDict.Keys
        recipientEmail = emailKey
        i = EmailsDict(recipientEmail)
        recipientName = ws.Cells(i, "S").Value

        Set OutMail = OutApp.CreateItem(0)
        With OutMail
            .Display ' This line is necessary to load the signature
            .To = recipientEmail
            .Subject = "Open Invoices"
            .HTMLBody = CreateEmailBody(ws, recipientEmail, recipientName) & .HTMLBody
            ' .Send ' Uncomment this line to send emails automatically
        End With
        Set OutMail = Nothing
        
        emailCounter = emailCounter + 1
        
        ' Send up to 29 emails a minute
        If emailCounter Mod 29 = 0 Then
            ' wait for 60 seconds before sending the next batch
            Application.Wait Now + TimeValue("00:01:00")
        End If
    Next

    Set OutApp = Nothing
End Sub

Function CreateEmailBody(ws As Worksheet, recipientEmail As String, recipientName As String) As String
    Dim i As Long
    Dim lastRow As Long
    Dim invoiceCount As Long
    Dim totalAmount As Double
    Dim mailBody As String

    lastRow = ws.Cells(ws.Rows.Count, "X").End(xlUp).Row
    invoiceCount = Application.WorksheetFunction.CountIf(ws.Range("X:X"), recipientEmail)
    totalAmount = Application.WorksheetFunction.SumIf(ws.Range("X:X"), recipientEmail, ws.Range("N:N"))

    mailBody = "Hello " & recipientName & "," & "<br>" & "<br>" & _
               "You have " & invoiceCount & " open invoices. That equates to " & Format(totalAmount, "$#,##0.00") & " dollars. " & vbNewLine & _
               "Please review and resolve to prevent the accounts from being placed on hold." & vbNewLine & vbNewLine & _
               "Invoice Details:" & "<br>" & "<br>" & _
               CreateInvoiceTable(ws, recipientEmail)

    CreateEmailBody = mailBody
End Function

Function CreateInvoiceTable(ws As Worksheet, recipientEmail As String) As String
    Dim i As Long
    Dim lastRow As Long
    Dim tableHTML As String

    lastRow = ws.Cells(ws.Rows.Count, "X").End(xlUp).Row
    tableHTML = "<table border='1' style='border-collapse:collapse'>" & _
                "<tr><th>Due Date</th><th>Vendor #</th><th>Vendor Name</th><th>Vendor Invoice #</th><th>PO Number</th>" & _
                "<th>IM User name</th><th>Invoice Subtotal</th><th>Invoice Tax</th><th>Invoice Total</th><th>Discrepency Status</th>" & _
                "<th>AP Clerk</th><th>PO Creator</th><th>CODA Username</th><th>PO Receipt Status</th></tr>"

    For i = 2 To lastRow
        If ws.Cells(i, "X").Value = recipientEmail Then
            tableHTML = tableHTML & "<tr>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "H").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "I").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "J").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "K").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "L").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "M").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "N").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "O").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "P").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "Q").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "R").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "S").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "T").Value & "</td>"
            tableHTML = tableHTML & "<td>" & ws.Cells(i, "U").Value & "</td>"
            tableHTML = tableHTML & "</tr>"
        End If
    Next i
    
    tableHTML = tableHTML & "</table>"
    
    CreateInvoiceTable = tableHTML

End Function
