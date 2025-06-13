Private Sub UserForm_Initialize()
    ' leave empty so nothing runs prematurely
End Sub

Public Sub PopulateForm()
    Set dictStoreData = CreateObject("Scripting.Dictionary")
    
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
        cmbStoreNumber.AddItem store(2)
    Next store
End Sub
