Sub Safety_47()
'Takes the Number of compound plates in a screening week , and returns a table of assay barcode values for csv export.

'To Do:
'Need to keep Assay list and output in seperate sheets
'Add Plate Type Column to output table
'Format output table headers


'Declare collections to store assay values and user input
Dim SFTY As New Collection
Dim Compound_Plates As New Collection
Dim Last_Row As Long
Dim New_Plate As Boolean

'Finding the last row with values
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row


'read in assay values from Assay list
For R = 1 To Last_Row

SFTY.Add Cells(R + 1, 1).Value 'assigning assay values to collection

Next R

'Request user for compound plate ID's
'Terminate when finished
Compound_Plates.Add InputBox("Please enter the compound plate ID", "Compound Plates") 'Ask the user for first ID

New_Plate = True 'initialize boolean check

While New_Plate = True


If MsgBox("Enter another Compound Plate ID?", vbYesNo) = vbYes Then 'prompts user to input another id if they clicked yes

Compound_Plates.Add InputBox("Please enter the compound plate ID", "Compound Plates")

Else

New_Plate = False ' if they click no then New_Plates value is changed to break the while loop

End If
Wend

'For each Compound Plate ID
'Concatenate the ID at the end of each assay value

Cells(1, 14).Value = "West_Label(Echo)" 'West barcode values
Cells(1, 15).Value = "South_Label(Text)" 'South barcode values
Row_Start = 2 ' Starting row of the output table

For ID = 1 To Compound_Plates.Count

For Assay = 1 To (SFTY.Count - 1)

Cells(Row_Start, 14).Value = SFTY(Assay) & "_" & Compound_Plates(ID)
Cells(Row_Start, 15).Value = SFTY(Assay) & "_" & Compound_Plates(ID)

Row_Start = Row_Start + 1

Next Assay
Next ID



End Sub
