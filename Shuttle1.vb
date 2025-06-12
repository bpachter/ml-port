''' At form level:

Dim dictStoreData As Object
Dim isInitialized As Boolean

Private Sub UserForm_Initialize()
    Set dictStoreData = CreateObject("Scripting.Dictionary")
    ' assume missingStores is passed in by calling macro
    Dim store As Variant
    For Each store In missingStores
        cmbStoreNumber.AddItem store
    Next
    cmbStoreNumber.List = SortArray(cmbStoreNumber.List)
    isInitialized = True
End Sub

Private Sub cmbStoreNumber_Change()
    If Not isInitialized Then Exit Sub
    
    ' save current store's values
    If cmbStoreNumber.Tag <> "" Then
        dictStoreData(cmbStoreNumber.Tag) = Array( _
            txtContractAccount.Text, txtSerialNumber.Text, txtBillingStart.Text, _
            txtBillingEnd.Text, txtBilledkWh.Text, txtBilledDemand.Text, _
            txtLoadFactor.Text, txtDemandKVar.Text)
    End If

    ' clear inputs
    txtContractAccount.Text = ""
    txtSerialNumber.Text = ""
    txtBillingStart.Text = ""
    txtBillingEnd.Text = ""
    txtBilledkWh.Text = ""
    txtBilledDemand.Text = ""
    txtLoadFactor.Text = ""
    txtDemandKVar.Text = ""

    ' load existing data if exists
    If dictStoreData.exists(cmbStoreNumber.Value) Then
        Dim values As Variant
        values = dictStoreData(cmbStoreNumber.Value)
        txtContractAccount.Text = values(0)
        txtSerialNumber.Text = values(1)
        txtBillingStart.Text = values(2)
        txtBillingEnd.Text = values(3)
        txtBilledkWh.Text = values(4)
        txtBilledDemand.Text = values(5)
        txtLoadFactor.Text = values(6)
        txtDemandKVar.Text = values(7)
    End If

    cmbStoreNumber.Tag = cmbStoreNumber.Value
End Sub

''' Continue button
Private Sub btnContinue_Click()
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
    Next

    ' write data to ASR - Bill Input
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("ASR - Bill Input")
    Dim startRow As Long: startRow = 6
    ws.Range("D6:L1000").ClearContents

    Dim r As Long: r = startRow
    For Each store In cmbStoreNumber.List
        vals = dictStoreData(store)
        ws.Cells(r, "D").Resize(1, 9).Value = vals
        r = r + 1
    Next

    Me.Hide
End Sub


''' Final Integration Steps
''' In RunBillingProcess:
If missingStores.Count > 0 Then
    Set formMissingStoreInput.missingStores = missingStores
    formMissingStoreInput.Show vbModeless
    Do While formMissingStoreInput.Visible
        DoEvents
    Loop
End If
