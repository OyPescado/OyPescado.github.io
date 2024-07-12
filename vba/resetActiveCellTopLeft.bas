Sub resetActiveCellTopLeft()

Dim currsheet As Worksheet
Dim sheet As Worksheet

Set currsheet = ActiveSheet

'Define the location to move the ActiveCell to
Const TopLeft As String = "A1"

'Loop through all the sheets in the workbook. Worksheet must be visible.
For Each sheet In Worksheets
    If sheet.Visible = xlSheetVisible Then Application.GoTo sheet.Range(TopLeft), scroll:=True
Next sheet

currsheet.Activate

End Sub