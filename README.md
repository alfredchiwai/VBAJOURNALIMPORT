
## Background:
### -A course-providing company which is using Square, a Retail POS system not directly integrated to MYOB.
### -Their course income may include both Current Month Revenue and Deferred Revenue, which are required to booked in the correct financial year.
### -Landlords require reporting monthly income for each shop, which necessitate monthly revenue to be booked in the correct month
### -Tons of manual works to turn Square Sales Report's csv file into the format with suitable details as per MYOB requirements & the Accounting Principle

## Purpose: 
### -Automated handling of Square Sales Report's csv file turning into required MYOB entries
### -As per Square Sales Report details, automatically dividing the sales into Current Month Revenue and Deferred Revenue and adding entries for them

## Benefits:
<img width="2072" height="673" alt="image" src="https://github.com/user-attachments/assets/16e007c0-4df9-4fcd-bb1b-c190f045c99e" />

## As-Is: Without automation, it requires tons of manual steps:

### Square POS Sales Report
<img width="2529" height="166" alt="螢幕擷取畫面 2025-10-28 125723" src="https://github.com/user-attachments/assets/d71e62b2-9fbd-409f-8a85-9494a197d1d3" />

### Manual Steps
<img width="2541" height="545" alt="螢幕擷取畫面 2025-10-28 125824" src="https://github.com/user-attachments/assets/874c8ade-df7e-48a8-af55-692479d416c6" />

<img width="2479" height="902" alt="螢幕擷取畫面 2025-10-28 125847" src="https://github.com/user-attachments/assets/19a49f6c-e864-4f2e-b5bc-4f58b6c71531" />

<img width="2513" height="599" alt="螢幕擷取畫面 2025-10-28 125908" src="https://github.com/user-attachments/assets/17ae87a8-588f-4818-b401-e2e1880c4b81" />

## With VBA Automation, Just a click away to turn it automatically!
[Please click to see the effects: https://youtu.be/JOgH-5jpioI](https://youtu.be/JOgH-5jpioI)

<details>
  <summary> VBA Code </summary>
  Sub SquareMYOB()


    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long
    Dim dateVal As String, journalNum As String, notes As String, customer As String, location As String
    Dim netTotal As Double, classCount As Long
    Dim className As String, dates() As String, monthYear As String
    Dim monthCounts As Object, monthKey As Variant
    Dim amountPerClass As Double, monthAmount As Double
    Dim outputFile As String, errorLog As String
    
    ' Set input and output worksheets
    Set wsInput = ActiveSheet
    Set wsOutput = ThisWorkbook.Sheets.Add
    wsOutput.Name = "MYOB_Import"
    
    ' MYOB Sales Invoice column headers
    With wsOutput
        .Cells(1, 1) = "Co./Last Name"
        .Cells(1, 2) = "Invoice #"
        .Cells(1, 3) = "Date"
        .Cells(1, 4) = "Description"
        .Cells(1, 5) = "Account #"
        .Cells(1, 6) = "Amount"
        .Cells(1, 7) = "Tax Code"
        .Cells(1, 8) = "Job"
        .Cells(1, 9) = "Memo"
    End With
    
    ' Create dictionary to count classes per month
    Set monthCounts = CreateObject("Scripting.Dictionary")
    
    ' Find last row in input sheet
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize output row and error log
    outputRow = 2
    errorLog = ""
    
    ' Loop through each row of input data (skip header)
    For i = 2 To lastRow
        ' Read input data
        dateVal = wsInput.Cells(i, 1).Value ' Date
        location = wsInput.Cells(i, 3).Value ' Location
        journalNum = wsInput.Cells(i, 4).Value ' Order ID
        customer = wsInput.Cells(i, 15).Value ' Customer Name
        notes = wsInput.Cells(i, 16).Value ' Notes
        netTotal = wsInput.Cells(i, 12).Value ' Net Total
        
        ' Skip empty rows
        If Len(dateVal) = 0 Or netTotal = 0 Then GoTo NextRow
        
        ' Parse Notes: Format "Class name - No. of Class: Date1, Date2, ..."
        Dim parts() As String
        parts = Split(notes, " - ")
        If UBound(parts) < 1 Then
            errorLog = errorLog & "Row " & i & ": Invalid Notes format - " & notes & vbCrLf
            GoTo NextRow
        End If
        
        className = Trim(parts(0)) ' e.g., "Lower Class"
        Dim subParts() As String
        subParts = Split(parts(1), ": ")
        If UBound(subParts) < 1 Then
            errorLog = errorLog & "Row " & i & ": Invalid Notes subformat - " & parts(1) & vbCrLf
            GoTo NextRow
        End If
        
        classCount = CLng(Replace(Split(subParts(0), " ")(0), " classes", "")) ' e.g., 8
        dates = Split(Trim(subParts(1)), ", ") ' e.g., "1/11/2024", "4/11/2024", ...
       
        ' Reset dictionary
        monthCounts.RemoveAll
        
        ' Count classes per month/year
        Dim j As Long
        For j = 0 To UBound(dates)
            Dim classDate As Variant
            classDate = ParseDate(Trim(dates(j))) ' Clean and parse date
            If IsDate(classDate) Then
                monthKey = Format(classDate, "mmm/yyyy")
                If Not monthCounts.exists(monthKey) Then
                    monthCounts(monthKey) = 1
                Else
                    monthCounts(monthKey) = monthCounts(monthKey) + 1
                End If
            Else
                errorLog = errorLog & "Row " & i & ": Invalid date in Notes - " & dates(j) & vbCrLf
            End If
        Next j
        
        ' Skip if no valid dates were found
        If monthCounts.Count = 0 Then GoTo NextRow
        
        ' Calculate amount per class
        amountPerClass = netTotal / classCount
        
        ' Write invoice lines for each month
        For Each monthKey In monthCounts.keys
            monthAmount = amountPerClass * monthCounts(monthKey)
            monthYear = monthKey
            
            With wsOutput
                .Cells(outputRow, 1) = customer ' Co./Last Name
                .Cells(outputRow, 2) = journalNum ' Invoice #
                .Cells(outputRow, 3) = dateVal ' Date
                .Cells(outputRow, 4) = className & " " & monthYear ' Description
                .Cells(outputRow, 5) = "1-1200" ' Accounts Receivable account
                .Cells(outputRow, 6) = Round(monthAmount, 2) ' Amount
                .Cells(outputRow, 7) = "N-T" ' Non-taxable, adjust if needed
                .Cells(outputRow, 8) = location ' Job
                .Cells(outputRow, 9) = customer & ": " & notes ' Memo
                outputRow = outputRow + 1
            End With
        Next monthKey
NextRow:
    Next i
    
    ' Save output as CSV
    outputFile = "myob_sales_invoices.csv"
    wsOutput.SaveAs Filename:=outputFile, FileFormat:=xlCSV, CreateBackup:=False
    
    ' Display error log if any issues
    If Len(errorLog) > 0 Then
        MsgBox "Processing complete with errors:" & vbCrLf & errorLog, vbExclamation
    Else
        MsgBox "MYOB Sales Invoices CSV saved as: " & outputFile
    End If
End Sub

' Custom function to parse DD/MM/YYYY dates
Function ParseDate(dateStr As String) As Variant
    On Error Resume Next
    Dim parts() As String
    parts = Split(Trim(dateStr), "/")
    If UBound(parts) = 2 Then
        ' Construct date as DD/MM/YYYY
        ParseDate = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
    Else
        ParseDate = Empty
    End If
    On Error GoTo 0
End Function

</details>
