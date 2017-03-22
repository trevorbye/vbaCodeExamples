Dim currentItem As String
Dim currentCustomer As String
Dim currentConcat As String
Dim findRow As Integer
Dim z As Integer
Dim lyRange As Range
Dim cyRange AS Range
Dim forecastRange As Range
Dim dashboardLY As Range
Dim dashboardCY As Range 
Dim dashboardForecast As Range

With Workbooks("Cheese Forecast Tool").Sheets("Dashboard")
	Set dashboardLY = .Range(.Cells(24,13), .Cells(24,64))
	Set dashboardCY = .Range(.Cells(25,13), .Cells(25,64))
	Set dashboardForecast = .Range(.Cells(26,13), .Cells(27,64))
End With


If ActiveSheet.Cells(2,3) = "All" Then
	'don't look up by customer; look by item
	If ActiveSheet.Cells(2,8) = "All" Then
	'return all items'
		With Workbooks("Cheese Forecast Tool").Sheets("Sheet1")
			Set lyRange = .Range(.Cells(63,6), .Cells(63,57))
			Set cyRange = .Range(.Cells(140,58), .Cells(140,109))
			Set forecastRange = .Range(.Cells(63,58), .Cells(63,109))
		End With

		dashboardLY.Value = lyRange.Value
		dashboardCY.Value = cyRange.Value
		dashboardForecast.Value = forecastRange.Value
	Else
		'return specific item'
		currentItem = Left((ActiveSheet.Cells(2,8).Value),6)

		'LY'
		With Workbooks("Cheese Forecast Tool").Sheets("Sheet1")
			For z = 52 To 62
				If .Cells(z,4).Value = currentItem Then
					Set lyRange = .Range(.Cells(z,6), .Cells(z,57))
					Set forecastRange = .Range(.Cells(z,58), .Cells(z,109))
					Exit For
				End If
			Next z

			For z = 129 To 139
				If .Cells(z, 56).Value = currentItem Then
					Set cyRange = .Range(.Cells(z, 58), .Cells(z, 109))
					Exit For
				End If 
			Next z
		End With
		'set ranges equal'
		dashboardLY.Value = lyRange.Value
		dashboardCY.Value = cyRange.Value
		dashboardForecast.Value = forecastRange.Value
	End If
Else 
	'looking by customer, but checking if it's a customer item combo, or the full customer amount
	currentCustomer = ActiveSheet.Cells(2,3).Value
	If Cells(2,8) = "All" Then
	'look up full amount for customer'
		With Workbooks("Cheese Forecast Tool").Sheets("Sheet1")
			For z = 67 To 75
				If .Cells(z, 5).Value = currentCustomer Then
					Set lyRange = .Range(.Cells(z,6), .Cells(z,57))
					Set forecastRange = .Range(.Cells(z,58), .Cells(z,109))
					Exit For
				End If
			Next z

			For z = 143 To 151
				If .Cells(z, 57).Value = currentCustomer Then
					Set cyRange = .Range(.Cells(z, 58), .Cells(z, 109))
					Exit For
				End If 
			Next z
		End With

		dashboardLY.Value = lyRange.Value
		dashboardCY.Value = cyRange.Value
		dashboardForecast.Value = forecastRange.Value

	Else
	'look up the specific customer/item combo'
	currentConcat = ActiveSheet.Cells(2,13).Value

		With Workbooks("Cheese Forecast Tool").Sheets("Sheet1")
			For z = 3 To 50
				If .Cells(z, 2).Value = currentConcat Then
					Set lyRange = .Range(.Cells(z,6), .Cells(z,57))
					Set forecastRange = .Range(.Cells(z,58), .Cells(z,109))
					Exit For
				End If
			Next z

			For z = 79 To 126
				If .Cells(z, 57).Value = currentConcat Then
					Set cyRange = .Range(.Cells(z, 58), .Cells(z, 109))
					Exit For
				End If 
			Next z
		End With
			On Error GoTo ErrMsg
			dashboardLY.Value = lyRange.Value
			Exit Sub
		dashboardCY.Value = cyRange.Value
		dashboardForecast.Value = forecastRange.Value
	End If
End If
ErrMsg:
MsgBox ("The selected customer/item combination does not exist.")