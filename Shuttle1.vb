Dim storeNum As Variant
Dim sortedStores As New Collection
Dim i As Long, j As Long

' sort the store numbers
For i = 1 To missingStores.Count
    Dim inserted As Boolean: inserted = False
    For j = 1 To sortedStores.Count
        If CLng(missingStores(i)(2)) < CLng(sortedStores(j)(2)) Then
            sortedStores.Add missingStores(i), , j
            inserted = True
            Exit For
        End If
    Next j
    If Not inserted Then sortedStores.Add missingStores(i)
Next i

' load into combo box
For Each storeNum In sortedStores
    Me.cboStoreNumber.AddItem storeNum(2) ' index 2 = Store Number
Next storeNum
