' add sorted store numbers to ComboBox
For Each store In sortedStores
    cmbStoreNumber.AddItem store(2)
Next store

' build full lookup from SerialTable worksheet
Dim wsLookup As Worksheet
Dim lastRow As Long
Dim storeVal As String, caVal As String, serialVal As String

Set wsLookup = ThisWorkbook.Sheets("Serial Number")
lastRow = wsLookup.Cells(wsLookup.Rows.Count, "A").End(xlUp).Row

For i = 2 To lastRow ' assuming row 1 has headers
    caVal = Trim(wsLookup.Cells(i, 1).Text)
    serialVal = Trim(wsLookup.Cells(i, 2).Text)
    storeVal = Trim(wsLookup.Cells(i, 3).Text)

    If Len(storeVal) > 0 Then
        dictStoreLookup(storeVal) = Array(caVal, serialVal)
        Debug.Print "[lookup initialized] store=" & storeVal & " CA=" & caVal & " SN=" & serialVal
    End If
Next i

Private Sub btnContinue_Click()
    ' save current store's values before validating
    If cmbStoreNumber.Tag <> "" Then
        dictStoreData(cmbStoreNumber.Tag) = Array( _
            txtContractAccount.Text, txtSerialNumber.Text, txtBillingStart.Text, _
            txtBillingEnd.Text, txtBilledkWh.Text, txtBilledDemand.Text, _
            txtLoadFactor.Text, txtDemandKVar.Text)
    End If

    ' validate all stores have data
    Dim store As Variant
    For Each store In cmbStoreNumber.List
        If Not dictStoreData.exists(store) Then
            MsgBox "Missing data for store " & store, vbExclamation
            Exit Sub
        End If

        Dim vals As Variant: vals = dictStoreData(store)
        Dim i As Long
        For i = 0 To UBound(vals)
            If Trim(vals(i)) = "" Then
                MsgBox "Incomplete entry for store " & store, vbExclamation
                Exit Sub
            End If
        Next i
    Next store

    ' write data to ASR - Bill Input
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("ASR - Bill Input")
    Dim startRow As Long: startRow = 6
    ws.Range("D6:L10000").ClearContents

    Dim r As Long: r = startRow
    For Each store In cmbStoreNumber.List
        vals = dictStoreData(store)
        ws.Cells(r, "D").Resize(1, 9).Value = vals
        r = r + 1
    Next store

    Me.Hide
End Sub
