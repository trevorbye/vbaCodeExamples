Dim lastDisplayRow As Integer
Dim firstDisplayRow As Integer
Dim currentItemNumber As Long
Dim orderItemNumber As Long
Dim orderQuantity As Long
Dim currentForecast As Long
Dim i As Integer
Dim z As Integer
Dim t As Integer
Dim j As Integer
Dim r As Integer
Dim currentPoundsPerCase As Integer
Dim requestedDeliveryDate As Long



'find last row in CustOrders'
With Workbooks("Cheese Production Plan.xlsm").Sheets("CustOrders")
    lastDisplayRow = .Cells(.Rows.Count, "C").End(xlUp).Row
    firstDisplayRow = 2
End With

'loop through each SKU in OrderTracker sheet'
For i = 4 To 14
    Dim arrayCritQuantities As New Collection

    With Workbooks("Cheese Production Plan.xlsm").Sheets("OrderTracker")
        currentItemNumber = .Cells(i, 1).Value
        currentForecast = .Cells(i, 4).Value
        currentPoundsPerCase = .Cells(i, 2).Value
    End With

    'loop through each order line'
    For z = firstDisplayRow To lastDisplayRow

        With Workbooks("Cheese Production Plan.xlsm").Sheets("CustOrders")
        orderItemNumber = .Cells(z, 19).Value
        orderQuantity = .Cells(z, 22).Value
        requestedDeliveryDate = .Cells(z, 13).Value
        End With

        If orderItemNumber = currentItemNumber Then
            'the critical quantity gets added to if order dates are within +- 3 days'
            Dim criticalOrderQuantity As Integer
            
            criticalOrderQuantity = orderQuantity * currentPoundsPerCase

            'loop through CustOrders again and compare dates to current order line'
            Dim secondaryOrderItemNumber As Long
            Dim secondaryRequestedDeliveryDate As Long
            Dim secondaryOrderQuantity As Long
            Dim daysBetween As Long

            For t = firstDisplayRow To lastDisplayRow
                With Workbooks("Cheese Production Plan.xlsm").Sheets("CustOrders")
                secondaryOrderItemNumber = .Cells(t, 19).Value
                secondaryOrderQuantity = .Cells(t, 22).Value
                secondaryRequestedDeliveryDate = .Cells(t, 13).Value
                End With

                If t = z Then
                    GoTo nullProcess2
                ElseIf secondaryOrderItemNumber = orderItemNumber Then
                    daysBetween = Abs(requestedDeliveryDate - secondaryRequestedDeliveryDate)
                    If daysBetween <= 2 Then
                    criticalOrderQuantity = criticalOrderQuantity + (secondaryOrderQuantity * currentPoundsPerCase)
                    End If
                End If
                
nullProcess2:
            Next t
            'store critical order quantity in collection
            Dim existsInCollection As Boolean
            Dim position As Integer
            existsInCollection = False
            
            For position = 1 To arrayCritQuantities.Count
                If arrayCritQuantities(position) = criticalOrderQuantity Then
                    existsInCollection = True
                    GoTo doesExist
                End If
            Next position
doesExist:
            If (existsInCollection = False) Then
            arrayCritQuantities.Add criticalOrderQuantity
            End If
        End If
    Next z

    'sort through collection to find largest critical quantity, populate to OrderTracker sheet'
    Dim trueCriticalQuantity As Long
    Dim currentCrit As Long
    Dim testCrit As Long
    trueCriticalQuantity = 0

    If arrayCritQuantities.Count = 1 Then
        trueCriticalQuantity = arrayCritQuantities(1)
    Else
        For j = 1 To arrayCritQuantities.Count
            currentCrit = arrayCritQuantities(j)
                For r = 1 To arrayCritQuantities.Count
                    testCrit = arrayCritQuantities(r)
                    If r = j Then
                        GoTo nullProcess
                    ElseIf currentCrit < testCrit Then
                        Exit For
                    ElseIf r = arrayCritQuantities.Count Then
                        trueCriticalQuantity = arrayCritQuantities(j)
                        Exit For
                    End If
nullProcess:
                Next r
    
                If trueCriticalQuantity <> 0 Then
                    Exit For
                End If
        Next j
    End If
    'populate to order tracker'
    With Workbooks("Cheese Production Plan.xlsm").Sheets("OrderTracker")
        .Cells(i, 5).Value = trueCriticalQuantity
    End With
    
    Set arrayCritQuantities = Nothing
Next i
