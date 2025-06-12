Public Function GetMissingStoresList() As Collection
    Dim wsGen As Worksheet
    Dim lastRow As Long, i As Long
    Dim store As String, serial As String, ca As String
    Dim missingStores As New Collection

    Set wsGen = ThisWorkbook.Sheets("Billing Interval Generator")
    lastRow = wsGen.Cells(wsGen.Rows.Count, "B").End(xlUp).Row

    ' loop from row 4 down (header is in row 3)
    For i = 4 To lastRow
        If Trim(wsGen.Cells(i, "H").Value) = "FALSE" Then
            store = Trim(wsGen.Cells(i, "B").Value)
            serial = Trim(wsGen.Cells(i, "G").Value)
            ca = ""  ' contract account will be filled in manually

            If Len(store) > 0 And Len(serial) > 0 Then
                missingStores.Add Array(ca, serial, store)
            End If
        End If
    Next i

    Set GetMissingStoresList = missingStores
End Function
