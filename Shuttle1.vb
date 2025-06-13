Private Sub cmbStoreNumber_Change()
    If Not isInitialized Then Exit Sub

    ' save current store's values
    If cmbStoreNumber.Tag <> "" Then
        dictStoreData(cmbStoreNumber.Tag) = Array( _
            txtContractAccount.Text, txtSerialNumber.Text, txtBillingStart.Text, _
            txtBillingEnd.Text, txtBilledkWh.Text, txtBilledDemand.Text, _
            txtLoadFactor.Text, txtDemandKVar.Text)
    End If

    ' clear all inputs
    txtContractAccount.Text = ""
    txtSerialNumber.Text = ""
    txtBillingStart.Text = ""
    txtBillingEnd.Text = ""
    txtBilledkWh.Text = ""
    txtBilledDemand.Text = ""
    txtLoadFactor.Text = ""
    txtDemandKVar.Text = ""

    ' preload any existing values
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
    Else
        ' fallback: fill contract account and serial number from original list
        Dim i As Long
        For i = 1 To pMissingStores.Count
            If CStr(pMissingStores(i)(2)) = CStr(cmbStoreNumber.Value) Then
                txtContractAccount.Text = CStr(pMissingStores(i)(0))
                txtSerialNumber.Text = CStr(pMissingStores(i)(1))
                Exit For
            End If
        Next i
    End If

    ' update tag for tracking
    cmbStoreNumber.Tag = cmbStoreNumber.Value
End Sub
