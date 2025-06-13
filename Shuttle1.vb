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

    ' build dictionary [serial] -> last END date
    Set dictLastEnd = CreateObject("Scripting.Dictionary")
    For i = 4 To lastRow
        serial = Trim(CStr(wsGen.Cells(i, "G").Value))
        If Len(serial) > 0 Then
            dictLastEnd(serial) = wsGen.Cells(i, "F").Value
        End If
    Next i

    ' build ASR dictionaries
    Set dictASROccurrences = CreateObject("Scripting.Dictionary")
    Set dictASREndDate = CreateObject("Scripting.Dictionary")

    Dim lastASRRow As Long: lastASRRow = wsASR.Cells(wsASR.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastASRRow
        serial = Trim(CStr(wsASR.Cells(i, "W").Value))
        billMonth = wsASR.Cells(i, "B").Value
        billYear = wsASR.Cells(i, "A").Value

        If serial <> "" And IsNumeric(billMonth) And IsNumeric(billYear) Then
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

                Debug.Print "[ASR FOUND] key=" & key
            End If
        End If
    Next i

    ' generate new rows
    Dim nextRow As Long: nextRow = lastRow + 1

    For i = 4 To lastRow
        store = CStr(wsGen.Cells(i, "B").Value)
        If wsGen.Cells(i, "C").Value = lastGenMonth And wsGen.Cells(i, "D").Value = lastGenYear Then
            serial = Trim(CStr(wsGen.Cells(i, "G").Value))

            If Len(store) > 0 And Len(serial) > 0 And dictLastEnd.exists(serial) Then
                key = serial & "|" & newMonth & "|" & newYear

                ' insert new row
                wsGen.Cells(nextRow, "B").Value = store ' only column B
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
                    Debug.Print "[Missing] key=" & key & ", store=" & store
                End If

                nextRow = nextRow + 1
            End If
        End If
    Next i

    MsgBox "Billing intervals for " & newMonth & "/" & newYear & " generated.", vbInformation
End Sub
