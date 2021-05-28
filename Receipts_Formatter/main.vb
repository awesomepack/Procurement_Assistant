Sub Receipt_Formatter()

'This sub routine formats coupa receipt export data into a format that can be used to
'generate pivot tables and pivot charts

' Defining my variables
Dim Last_Row As Long, Counter As Integer
Dim Total_Ordered As Integer, Items() As String



' Assigning values
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row ' Last row in the data
Counter = 2 ' Starting row of summary table
Start_Date = Cells(2, 4).Value


'Determining the duration of a blanket order
Start_Month = Left(Cells(2, 4).Value, 2) 'Start month of blanket order
End_Month = Left(Cells(2, 5).Value, 2) ' End month of blanket order
Duration = (End_Month - Start_Month) + 1

'Extracting the total items ordered on the blanket order

'split the Giant string in "Items" using split
Items = Split(Cells(28, 7).Value, ",")

For c = 1 To Len(Items(1))

MsgBox (Mid(Items(1), c, 1))

Next c



'calculated values
Total_Ordered = 110 'extract this number from "Items"
Shipment_Qty = Total_Ordered / Duration

'populating summary table with duration rows
While Duration <> 0

Cells(Counter, 13).Value = Cells(2, 6).Value
Cells(Counter, 14).Value = Total_Ordered
Cells(Counter, 15).Value = Start_Date

Start_Date = DateAdd("M", 1, Start_Date)
Counter = Counter + 1
Duration = Duration - 1
Wend