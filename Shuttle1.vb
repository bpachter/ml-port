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
