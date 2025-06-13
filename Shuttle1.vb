Private Sub UserForm_Initialize()
    Set dictStoreData = CreateObject("Scripting.Dictionary")
    
    Dim store As Variant
    For Each store In pMissingStores
        ' store = Array(ca, serial, storeNumber)
        cmbStoreNumber.AddItem store(2) ' index 2 = Store Number
    Next store

    ' sort the store numbers
    Dim sortedStores As New Collection
    Dim i As Long, j As Long
    Dim inserted As Boolean

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

    ' clear existing combo
    cmbStoreNumber.Clear

    ' populate combo with sorted store numbers
    For i = 1 To sortedStores.Count
        cmbStoreNumber.AddItem sortedStores(i)(2)
    Next i

    isInitialized = True
End Sub



' step 3: in RunBillingProcess
' step 3: identify missing stores and launch interactive form
Set missingStores = GetMissingStoresList()
If missingStores.Count > 0 Then
    Set formMissingStoreInput.MissingStores = missingStores
    formMissingStoreInput.Show vbModeless
    Do While formMissingStoreInput.Visible
        DoEvents
    Loop
End If
