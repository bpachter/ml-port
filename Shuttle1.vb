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
                sortedStores.Add pMissingStores(i), , j
                inserted = True
                Exit For
            End If
        Next j
        If Not inserted Then sortedStores.Add pMissingStores(i)
    Next i

    cmbStoreNumber.Clear
    For Each store In sortedStores
        cmbStoreNumber.AddItem store(2) ' index 2 = Store Number
        dictStoreLookup(store(2)) = Array(store(0), store(1)) ' (ContractAccount, SerialNumber)
    Next store

    isInitialized = True
End Sub
