Public Sub PopulateForm()
    Set dictStoreData = CreateObject("Scripting.Dictionary")
    Set dictStoreLookup = CreateObject("Scripting.Dictionary")

    Dim sortedStores As New Collection
    Dim i As Long, j As Long, inserted As Boolean
    Dim store As Variant

    ' sort pMissingStores by store number (element 3)
    For i = 1 To pMissingStores.Count
        inserted = False
        For j = 1 To sortedStores.Count
            If CLng(pMissingStores(i)(2)) < CLng(sortedStores(j)(2)) Then
                sortedStores.Add pMissingStores(i), before:=j
                inserted = True
                Exit For
            End If
        Next j
        If Not inserted Then sortedStores.Add pMissingStores(i)
    Next i

    ' clear and populate combo box
    cmbStoreNumber.Clear
    For Each store In sortedStores
        cmbStoreNumber.AddItem store(2) ' store number
        ' explicitly store as string array to prevent type mismatch
        dictStoreLookup(CStr(store(2))) = Array(CStr(store(0)), CStr(store(1))) ' (CA, Serial)
    Next store

    isInitialized = True
End Sub

 


 Private Sub cmbStoreNumber_Change()
    If Not isInitialized Then Exit Sub

    ' save current values before switching
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

    ' try to load previous entry
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

        Debug.Print "[loaded from dictStoreData] ContractAccount=" & txtContractAccount.Text
    Else
        ' fallback: load from SerialTable lookup dictionary
        Dim storeKey As String
        storeKey = CStr(cmbStoreNumber.Value)

        Debug.Print "cmbStoreNumber.Value = [" & storeKey & "]"
        If dictStoreLookup.exists(storeKey) Then
            Dim arr As Variant
            arr = dictStoreLookup(storeKey)
            txtContractAccount.Text = arr(0)
            txtSerialNumber.Text = arr(1)

            Debug.Print "[lookup fallback] CA=" & arr(0) & " SN=" & arr(1)
        Else
            Debug.Print "[lookup FAILED] dictStoreLookup does not contain: " & storeKey
        End If
    End If

    ' store current selection for later save
    cmbStoreNumber.Tag = cmbStoreNumber.Value
End Sub
