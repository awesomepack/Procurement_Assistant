Sub Receipt_Formatter()

'This sub routine formats coupa receipt export data into a format that can be used to
'generate pivot tables and pivot charts

' Defining Variable data types
Dim Last_Row As Long, Counter As Integer
Dim Total_Ordered As Variant, Items() As String

'Headers for the summary table
Cells(1, 12).Value = "Consumable ID"
Cells(1, 13).Value = "PO"
Cells(1, 14).Value = "Order_Qty"
Cells(1, 15).Value = "Scheduled_Delivery"
Cells(1, 16).Value = "Shipment_Qty"
Cells(1, 17).Value = "Actual_Delivery"
Cells(1, 18).Value = "Received_qty"

Last_Row = Cells(Rows.Count, 1).End(xlUp).Row ' Last row in the data
Counter = 2 ' Starting row of summary table


'Start For Loop

'For Row = 2 To Last_Row


' Assigning values

Start_Date = Cells(34, 4).Value


'Determining the duration of a blanket order
Start_Month = Int(Left(Cells(34, 4).Value, 2)) 'Start month of blanket order
End_Month = Left(Cells(34, 5).Value, 2) ' End month of blanket order
Duration = Abs(Int((End_Month - Start_Month))) + 1 'Fix_1:apply absolute function to difference

'Extracting the total items ordered on the blanket order
'split "Items" by individual line Item
Items = Split(Cells(34, 7).Value, ",")

'looping through the first 4 characters of each line Item
Curr_Str = ""
For c = 1 To 4
Curr_Char = Mid(Items(0), c, 1)

If IsNumeric(Curr_Char) Then

Curr_Str = Curr_Str & Curr_Char

End If
Next c

If IsNumeric(Curr_Str) Then

'calculating shipment qty from Total order qty and duration
Total_Ordered = Int(Curr_Str)
Shipment_Qty = Total_Ordered / Duration

Else

Total_Ordered = "Amt Based Order"
Shipment_Qty = "Amt"

End If


'populating summary table with duration rows
'issue_1: Negative Duration causes infinite loop
While Duration <> 0

Cells(Counter, 13).Value = Cells(34, 6).Value
Cells(Counter, 14).Value = Total_Ordered
Cells(Counter, 15).Value = Start_Date
Cells(Counter, 16).Value = Shipment_Qty

Start_Date = DateAdd("M", 1, Start_Date)
Counter = Counter + 1
Duration = Duration - 1
Wend

'Next Row
    
'End For loop


End Sub
