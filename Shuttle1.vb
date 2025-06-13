Private Sub btnContinue_Click()
    ' force save the currently selected store's inputs
    If cmbStoreNumber.value <> "" Then
        dictStoreData(CStr(cmbStoreNumber.value)) = Array( _
            txtContractAccount.Text, txtSerialNumber.Text, txtBillingStart.Text, _
            txtBillingEnd.Text, txtBilledkWh.Text, txtBilledDemand.Text, _
            txtLoadFactor.Text, txtDemandKVar.Text _
        )
    End If

    ' validate all stores have complete data
    Dim store As Variant
    For Each store In cmbStoreNumber.List
        If Not IsNull(store) And Trim(store) <> "" Then
            Dim key As String: key = CStr(store)

            If Not dictStoreData.exists(key) Then
                MsgBox "Missing data for store " & store, vbExclamation
                Exit Sub
            End If

            Dim vals As Variant: vals = dictStoreData(key)
            Dim i As Long
            For i = 0 To UBound(vals)
                If Trim(vals(i)) = "" Then
                    MsgBox "Incomplete entry for store " & store, vbExclamation
                    Exit Sub
                End If
            Next i
        End If
    Next store

    ' write data to ASR - Bill Input sheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("ASR - Bill Input")
    Dim startRow As Long: startRow = 6
    ws.Range("D6:L10000").ClearContents

    Dim r As Long: r = startRow
    For Each store In cmbStoreNumber.List
        If Not IsNull(store) And Trim(store) <> "" Then
            Dim values As Variant: values = dictStoreData(CStr(store))
            ws.Cells(r, "D").Resize(1, 9).value = values
            r = r + 1
        End If
    Next store

    Me.Hide
End Sub

Public Sub GenerateNextBillingInterval()
    Dim wsGen As Worksheet, wsASR As Worksheet
    Dim i As Long, lastRow As Long, newMonth As Long, newYear As Long
    Dim store As String, serial As String
    Dim startDate As Date, endDate As Date
    Dim dictLastEnd As Object, dictASROccurrences As Object, dictASREndDate As Object
    Dim billMonth As Variant, billYear As Variant
    Dim lastGenMonth As Long, lastGenYear As Long
    Dim key As String

    Set wsGen = ThisWorkbook.Sheets("Billing Interval Generator")
    Set wsASR = ThisWorkbook.Sheets("ASR")

    ' delete rows without "IST-2" in column L of ASR
    lastRow = wsASR.Cells(wsASR.Rows.Count, "D").End(xlUp).Row
    For i = lastRow To 2 Step -1
        If wsASR.Cells(i, "L").Value <> "IST-2" Then
            wsASR.Rows(i).Delete
        End If
    Next i

    ' determine last filled month and year
    lastRow = wsGen.Cells(wsGen.Rows.Count, "B").End(xlUp).Row
    If lastRow < 4 Then
        MsgBox "No billing interval data found.", vbCritical
        Exit Sub
    End If

    lastGenMonth = wsGen.Cells(lastRow, "C").Value
    lastGenYear = wsGen.Cells(lastRow, "D").Value

    ' advance to next month
    newMonth = lastGenMonth + 1
    newYear = lastGenYear
    If newMonth > 12 Then
        newMonth = 1
        newYear = newYear + 1
    End If

    ' build dictionary [serial] -> last END date from wsGen
    Set dictLastEnd = CreateObject("Scripting.Dictionary")
    For i = 4 To lastRow
        serial = Trim(wsGen.Cells(i, "G").Value)
        If Len(serial) > 0 Then
            dictLastEnd(serial) = wsGen.Cells(i, "F").Value
        End If
    Next i

    ' build ASR dictionaries
    Set dictASROccurrences = CreateObject("Scripting.Dictionary")
    Set dictASREndDate = CreateObject("Scripting.Dictionary")

    Dim lastASRRow As Long: lastASRRow = wsASR.Cells(wsASR.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastASRRow
        serial = Trim(wsASR.Cells(i, "W").Value)
        billMonth = wsASR.Cells(i, "B").Value
        billYear = wsASR.Cells(i, "A").Value

        If serial <> "" Then
            If IsNumeric(billMonth) And IsNumeric(billYear) Then
                If CLng(billMonth) = newMonth And CLng(billYear) = newYear Then
                    key = serial & "|" & newMonth & "|" & newYear
                    If dictASROccurrences.exists(key) Then
                        dictASROccurrences(key) = dictASROccurrences(key) + 1
                    Else
                        dictASROccurrences(key) = 1
                    End If

                    If IsDate(wsASR.Cells(i, "P").Value) Then
                        dictASREndDate(key) = wsASR.Cells(i, "P").Value
                    End If
                End If
            Else
                Debug.Print "[SKIPPED] serial=" & serial & _
                            ", billMonth=" & billMonth & _
                            ", billYear=" & billYear & _
                            ", row=" & i
            End If
        End If
    Next i

    ' generate new rows starting after lastRow
    Dim nextRow As Long: nextRow = lastRow + 1

    For i = 4 To lastRow
        store = wsGen.Cells(i, "B").Value
        If wsGen.Cells(i, "C").Value = lastGenMonth And wsGen.Cells(i, "D").Value = lastGenYear Then
            serial = Trim(wsGen.Cells(i, "G").Value)

            If Len(store) > 0 And Len(serial) > 0 And dictLastEnd.exists(serial) Then
                key = serial & "|" & newMonth & "|" & newYear

                ' insert new row
                wsGen.Cells(nextRow, "A").Value = store
                wsGen.Cells(nextRow, "B").Value = store
                wsGen.Cells(nextRow, "C").Value = newMonth
                wsGen.Cells(nextRow, "D").Value = newYear
                wsGen.Cells(nextRow, "E").Value = dictLastEnd(serial) + 1
                wsGen.Cells(nextRow, "G").Value = serial

                If dictASREndDate.exists(key) Then
                    wsGen.Cells(nextRow, "F").Value = dictASREndDate(key)
                Else
                    wsGen.Cells(nextRow, "F").Value = ""
                End If

                If dictASROccurrences.exists(key) Then
                    wsGen.Cells(nextRow, "H").Value = True
                    wsGen.Cells(nextRow, "I").Value = dictASROccurrences(key)
                    wsGen.Cells(nextRow, "J").Value = ""
                Else
                    wsGen.Cells(nextRow, "H").Value = False
                    wsGen.Cells(nextRow, "I").Value = 0
                    wsGen.Cells(nextRow, "J").Value = "BILL INPUT"
                End If

                nextRow = nextRow + 1
            End If
        End If
    Next i

    MsgBox "Billing intervals for " & newMonth & "/" & newYear & " generated.", vbInformation
End Sub
