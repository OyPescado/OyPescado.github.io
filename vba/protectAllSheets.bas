Sub ProtectAll()
    Dim wsh As Worksheet
    For Each wsh In Worksheets
        wsh.Protect Password:="Secret"
    Next wsh
End Sub