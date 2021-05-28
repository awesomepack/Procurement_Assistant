Sub Safety_47()
'Takes the Number of compound plates in a screening week , and returns a table of assay barcode values for csv export.


'Declare Array to hold values for 72 unique Assays
Dim SFTY(1 To 72) As String
Dim Compound_Plate(1 To 10) As String

'read in assay values from Assay list
'For R = 1 To UBound(SFTY)

'SFTY(R) = Cells(R + 1, 1).Value 'assigning assay values to array

'Next R



'Request for compound plate value until all values have been entered

Other_Plate = True
Plate_Count = 1

While Other_Plate = True

Compound_Plate(Plate_Count) = InputBox("Please enter the compound plate ID")
MsgBox (Compound_Plate(Plate_Count))

Response = InputBox("Enter another compound plate ID [y/n]?")

If Response = "n" Then

Other_Plate = False

End If

Plate_Count = Plate_Count + 1
Wend









End Sub
