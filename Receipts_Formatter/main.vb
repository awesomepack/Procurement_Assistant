Sub Receipt_Formatter()

'To Do:
'


'This sub routine formats coupa receipt export data into a format that can be used to
'generate pivot tables and pivot charts

' Defining Variable data types
Dim Last_Row As Long
Dim Result_Table_Counter As Integer

Result_Table_Counter = 2 'starting row of summary table

'The header rows for "Shipments_Per_Order" sheet

'Headers for the summary table
With Sheets(4)
Sheets(4).Activate

Cells(1, 1).Value = "Consumable ID"
Cells(1, 2).Value = "PO"
Cells(1, 3).Value = "Order_Qty"
Cells(1, 4).Value = "Scheduled_Delivery"
Cells(1, 5).Value = "Shipment_Qty"
Cells(1, 6).Value = "Actual_Delivery"
Cells(1, 7).Value = "Received_qty"

End With


'Collecting the data from the "Spotfire_Output" sheet
    
    
Last_Row = Sheets(3).Cells(Rows.Count, 1).End(xlUp).Row()
For Row = 2 To Last_Row

With Sheets(3)
Sheets(3).Activate
    
Consumable_ID = Sheets(3).Cells(Row, 1).Value
Purchase_Order = Sheets(3).Cells(Row, 2).Value
Duration = Sheets(3).Cells(Row, 5).Value
Order_Qty = Sheets(3).Cells(Row, 7).Value
Shipment_Qty = Sheets(3).Cells(Row, 8).Value
Start_Date = Sheets(3).Cells(Row, 3).Value
End With


While Duration <> 0

With Sheets(4)
Sheets(4).Activate

Cells(Result_Table_Counter, 1).Value = Consumable_ID
Cells(Result_Table_Counter, 2).Value = Purchase_Order
Cells(Result_Table_Counter, 3).Value = Order_Qty
Cells(Result_Table_Counter, 4).Value = Start_Date
Cells(Result_Table_Counter, 5).Value = Shipment_Qty

Start_Date = DateAdd("m", 1, Start_Date)
Result_Table_Counter = Result_Table_Counter + 1
Duration = Duration - 1
End With


Wend
Next Row


    
End Sub

