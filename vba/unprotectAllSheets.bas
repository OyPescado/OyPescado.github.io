Sub UnprotectAll()
    Dim wsh As Worksheet
    For Each wsh In Worksheets
        wsh.Unprotect Password:="Secret"
    Next wsh
End Sub