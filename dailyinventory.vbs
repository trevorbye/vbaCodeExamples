Private Sub CommandButton1_Click()

Dim wkb As Workbook
Dim lastRow As Integer
Dim lastColumn As Integer
Dim i, t, j, z, r, k, w, f, u, e, d, v, n, p, b, aa As Integer
Dim yesCount As Integer
Dim finalArrayCount As Integer
Dim lastDBRow As Integer
Dim lastMacroRow As Long
Dim verylastDBRow As Integer
Dim bookName As String
Dim bookDate As String
Dim dateString As String
Dim activePaste As String
Dim matchDate As String
Dim txt As String
Dim length As Long
Dim startColumn As Long
startColumn = (Application.ActiveWorkbook.Sheets("Database(CU's)").Cells(3, Columns.Count).End(xlToLeft).Column) + 1
Dim bookCount As Integer
bookCount = Application.Workbooks.Count - 2
Dim wkbArray() As String
Dim duplicateArray() As Variant
Dim finalArray() As Variant
ReDim wkbArray((bookCount - 1), 1) As String

'Loop through each workbook, store book name and date from X2 in a 2d array'

Application.ActiveWorkbook.Sheets("macroPaste").Visible = True

i = 0
For Each wkb In Workbooks
    If Left(wkb.Name, 15) = "CP_Inventory_By" Then

        dateString = wkb.ActiveSheet.Range("X2").Value
        bookName = wkb.Name
        length = Len(dateString)
        
        'Format string based on total string length
        If length = 19 Then
            bookDate = Left(dateString, 8)
        ElseIf length = 20 Then
            bookDate = Left(dateString, 9)
        ElseIf length = 21 Then
            bookDate = Left(dateString, 10)
        End If

        'Add book name and date to array'

        wkbArray(i, 0) = bookName
        wkbArray(i, 1) = bookDate
        i = i + 1
    Else
    End If
Next wkb


'create loop to specify number of times to run paste operation'

For t = 1 To bookCount
    matchDate = Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("Database(CU's)").Cells(1, startColumn).Value

        'Find book name based on match date'
        d = 0
        n = 0
        For j = LBound(wkbArray) To UBound(wkbArray)
            If wkbArray(d, 1) = matchDate Then
            n = n + d
            Exit For
            End If
            d = d + 1
        Next j
        
        activePaste = wkbArray(n, 0)
        With Workbooks(activePaste).Sheets("CP_Inventory_By_Run_Date_Email")
            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With

        'Set macroPaste Range equal to activePaste range, filter criteria.'

        Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("macroPaste").Range(Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("macroPaste").Cells(1, 1), Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("macroPaste").Cells(lastRow, 24)).Value = Workbooks(activePaste).Sheets("CP_Inventory_By_Run_Date_Email").Range(Workbooks(activePaste).Sheets("CP_Inventory_By_Run_Date_Email").Cells(1, 1), Workbooks(activePaste).Sheets("CP_Inventory_By_Run_Date_Email").Cells(lastRow, 24)).Value

        With Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("macroPaste")
            lastMacroRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            .Range(.Cells(1, 1), .Cells(lastMacroRow, 24)).AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=.Range("AA1:AA12"), Unique:=False
            .UsedRange.Copy
        End With

        'Paste in daily paste sheet,

        With Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("Paste Daily Data")
            .Range("E1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            currentLastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
            yesCount = Application.WorksheetFunction.CountIf(.Range(.Cells(2, 3), .Cells(currentLastRow, 3)), "Yes")
        End With



        'Create Array of "YES Database Items'
        If yesCount > 0 Then
            With Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("Paste Daily Data")

                ReDim duplicateArray(yesCount, 2) As Variant
                r = 0

                For z = 2 To currentLastRow
                    If .Cells(z, 3).Value = "Yes" Then
                        duplicateArray(r, 0) = .Cells(z, 5).Value
                        duplicateArray(r, 1) = .Cells(z, 6).Value
                        duplicateArray(r, 2) = .Cells(z, 9).Value
                        r = r + 1
                    Else
                    End If
                Next z
            End With

            'Create final array with unique YES items'
            ReDim finalArray(yesCount, 2) As Variant
            finalArrayCount = 0
            k = 0
            f = 0
            'Figure our how many times to loop through duplicate array'
            p = 0
            For k = LBound(duplicateArray) To UBound(duplicateArray)
                'Figure out if the value is already in the final array'
                v = 0
                For f = LBound(finalArray) To UBound(finalArray)
                    If finalArray(f, 1) = duplicateArray(k, 1) Then
                    v = v + 1
                    Exit For
                    End If
                Next f
                'if the value isn't in the final array, then add it. Otherwise, next k
                If v <> 1 Then
                    finalArray(p, 1) = duplicateArray(p, 1)
                    finalArray(p, 0) = duplicateArray(p, 0)
                    finalArray(p, 2) = duplicateArray(p, 2)
                    finalArrayCount = finalArrayCount + 1
                    p = p + 1
                End If
                
            Next k

            'Add new values from finalArray to bottom of DatabaseCU sheet'
            e = 0
            b = 0
            With Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("Database(CU's)")
                lastDBRow = (.Cells(.Rows.Count, "D").End(xlUp).Row) + 1
                    For e = LBound(finalArray) To UBound(finalArray)
                        .Cells(lastDBRow, 2).Value = finalArray(e, 0)
                        .Cells(lastDBRow, 3).Value = finalArray(e, 1)
                        .Cells(lastDBRow, 4).Value = finalArray(e, 2)
                        lastDBRow = lastDBRow + 1
                    Next e
            End With
        End If

        'fill down formula and move to next sheet'


        With Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("Database(CU's)")
            verylastDBRow = .Cells(.Rows.Count, "D").End(xlUp).Row
            .Range(.Cells(2, startColumn), .Cells(2, startColumn)).AutoFill Destination:=.Range(.Cells(2, startColumn), .Cells(verylastDBRow, startColumn)), Type:=xlFillDefault
            .Range(.Cells(2, startColumn), .Cells(verylastDBRow, startColumn)).Copy
            .Range(.Cells(2, startColumn), .Cells(verylastDBRow, startColumn)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
        
        'Clear daily paste
        With Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("Paste Daily Data")
            .Range(.Cells(2, 5), .Cells(currentLastRow, 28)).Clear
        End With
        
        'clear macro paste
        With Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("macroPaste")
            .Range(.Cells(1, 1), .Cells(lastMacroRow, 24)).Clear
            On Error Resume Next
            .ShowAllData
            On Error GoTo 0
        End With
        
        'Erase Arrays
        Erase finalArray, duplicateArray

        startColumn = startColumn + 1
Next t

Workbooks("CP Inventory Metrics with Pallets new.xlsm").Sheets("macroPaste").Visible = False
MsgBox "Script finished."
End Sub
