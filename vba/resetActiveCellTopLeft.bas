Dim currsheet As Worksheet
Dim sheet As Worksheet
Set currsheet = ActiveSheet
'Change A1 to suit your preference
Const TopLeft As String = "A1"
'Loop through all the sheets in the workbook
For Each sheet In Worksheets
    'Only does this for visible worksheets
    If sheet.Visible = xlSheetVisible Then Application.GoTo sheet.Range(TopLeft), scroll:=True
Next sheet
currsheet.Activate