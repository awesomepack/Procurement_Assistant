Sub Safety_47()
'Takes the Number of compound plates in a screening week , and returns a table of assay barcode values for csv export.


'Declare Array to hold values for 72 unique Assays
Dim SFTY(1 To 72) As String

'read in assay values from Assay list
For R = 1 To UBound(SFTY)

SFTY(R) = Cells(R + 1, 1).Value 'assigning assay values to array

Next R





End Sub
