Public Sub PopulateForm()
    Dim wsSerial As Worksheet
    Dim serialData As Variant
    Dim i As Long, j As Long
    Dim inserted As Boolean
    Dim store As Variant
    Dim sortedStores As New Collection

    ' init store data dictionaries
    Set dictStoreData = CreateObject("Scripting.Dictionary")
    Set dictStoreLookup = CreateObject("Scripting.Dictionary")

    ' load SerialTable from Serial Number worksheet
    Set wsSerial = ThisWorkbook.Sheets("Serial Number")
    serialData = wsSerial.Range("SerialTable").Value

    ' build lookup dictionary: store number → (contract account, serial number)
    For i = 2 To UBound(serialData, 1) ' skip header row
        If Not dictStoreLookup.exists(CStr(serialData(i, 3))) Then ' col 3 = store number
            dictStoreLookup.Add CStr(serialData(i, 3)), Array(CStr(serialData(i, 1)), CStr(serialData(i, 2))) ' contract, serial
        End If
    Next i

    ' sort pMissingStores by store number (element 3)
    For i = 1 To pMissingStores.Count
        inserted = False
        For j = 1 To sortedStores.Count
            If CLng(pMissingStores(i)(2)) < CLng(sortedStores(j)(2)) Then
                sortedStores.Add pMissingStores(i), , j
                inserted = True
                Exit For
            End If
        Next j
        If Not inserted Then sortedStores.Add pMissingStores(i)
    Next i

    ' populate combo box with sorted store numbers
    cmbStoreNumber.Clear
    For Each store In sortedStores
        cmbStoreNumber.AddItem store(2)
    Next store

    isInitialized = True
End Sub
 


 Private Sub cmbStoreNumber_Change()
    If Not isInitialized Then Exit Sub

    ' save current store's values before switching
    If cmbStoreNumber.Tag <> "" Then
        dictStoreData(cmbStoreNumber.Tag) = Array( _
            txtContractAccount.Text, txtSerialNumber.Text, txtBillingStart.Text, _
            txtBillingEnd.Text, txtBilledkWh.Text, txtBilledDemand.Text, _
            txtLoadFactor.Text, txtDemandKVar.Text _
        )
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

    ' preload if user already entered data
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
        ' fallback: use lookup from SerialTable
        If dictStoreLookup.exists(CStr(cmbStoreNumber.Value)) Then
            Dim arr() As String
            arr = dictStoreLookup(CStr(cmbStoreNumber.Value))
            txtContractAccount.Text = arr(0)
            txtSerialNumber.Text = arr(1)
        End If
    End If

    ' tag used to track current store context
    cmbStoreNumber.Tag = cmbStoreNumber.Value
End Sub
