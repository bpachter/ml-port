Public Function GetMissingStoresList() As Collection
    Dim wsGen As Worksheet
    Dim lastRow As Long, i As Long
    Dim maxMonth As Long, maxYear As Long
    Dim thisMonth As Long, thisYear As Long
    Dim store As String, serial As String, ca As String
    Dim missingStores As New Collection

    Set wsGen = ThisWorkbook.Sheets("Billing Interval Generator")
    lastRow = wsGen.Cells(wsGen.Rows.Count, "B").End(xlUp).Row

    ' step 1: find the latest month and year combination
    For i = 4 To lastRow
        thisMonth = Val(wsGen.Cells(i, "C").Value)
        thisYear = Val(wsGen.Cells(i, "D").Value)
        If thisYear > maxYear Or (thisYear = maxYear And thisMonth > maxMonth) Then
            maxMonth = thisMonth
            maxYear = thisYear
        End If
    Next i

    ' step 2: collect missing stores only for latest month/year
    For i = 4 To lastRow
        thisMonth = Val(wsGen.Cells(i, "C").Value)
        thisYear = Val(wsGen.Cells(i, "D").Value)

        If thisMonth = maxMonth And thisYear = maxYear Then
            If wsGen.Cells(i, "H").Value = False Then
                store = Trim(wsGen.Cells(i, "B").Value)
                serial = Trim(wsGen.Cells(i, "G").Value)
                ca = "" ' contract account to be filled by user
                If Len(store) > 0 And Len(serial) > 0 Then
                    missingStores.Add Array(ca, serial, store)
                End If
            End If
        End If
    Next i

    Set GetMissingStoresList = missingStores
End Function
