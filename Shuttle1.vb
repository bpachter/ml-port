Private Sub btnContinue_Click()
    ' force save the currently selected store's inputs
    If cmbStoreNumber.value <> "" Then
        dictStoreData(CStr(cmbStoreNumber.value)) = Array( _
            txtContractAccount.Text, txtSerialNumber.Text, txtBillingStart.Text, _
            txtBillingEnd.Text, txtBilledkWh.Text, txtBilledDemand.Text, _
            txtLoadFactor.Text, txtDemandkVar.Text _
        )
    End If

    ' validate all stores have complete data
    Dim store As Variant
    For Each store In cmbStoreNumber.List
        Dim key As String: key = CStr(store)

        If Not dictStoreData.exists(key) Then
            MsgBox "Missing data for store " & store, vbExclamation
            Exit Sub
        End If

        Dim vals As Variant
        vals = dictStoreData(key)
        Dim i As Long
        For i = 0 To UBound(vals)
            If Trim(vals(i) & "") = "" Then
                MsgBox "Incomplete entry for store " & store, vbExclamation
                Exit Sub
            End If
        Next i
    Next store

    ' write data to ASR - Bill Input sheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("ASR - Bill Input")
    Dim startRow As Long: startRow = 6
    ws.Range("D6:L10000").ClearContents

    Dim r As Long: r = startRow
    For Each store In cmbStoreNumber.List
        vals = dictStoreData(CStr(store))
        ws.Cells(r, "D").Resize(1, 9).value = vals
        r = r + 1
    Next store

    Me.Hide
End Sub
