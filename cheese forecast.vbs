Const olFolderInbox As Integer = 6
Const AttachmentPath As String = "C:\Users\TRBYE\Desktop\cheeseReport.csv"
Const AttachmentPath2 As String = "C:\Users\TRBYE\Desktop\boiseExtract.csv"

Private Sub CommandButton1_Click()

Application.ScreenUpdating = False
Dim oOlAp As Object, oOlns As Object, oOlInb As Object, oOlItm As Object, oOltargetEmail As Object, oOlAtch As Object
Dim beginningDate As String, endingDate As String, todaysDateTime As String, todaysDate As String, receivedTime As String, receivedTime2 As String, date1 As String
Dim x As Integer
Dim emailFileName As String

Dim receivedTimeBoise As String
Dim receivedTimeBoise2 As String
Dim beginningDateBoise As String
Dim endingDateBoise As String



Set oOlAp = GetObject(, "Outlook.application")
Set oOlns = oOlAp.GetNamespace("MAPI")
Set oOlInb = oOlns.GetDefaultFolder(olFolderInbox)

Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic

receivedTime = " 06:08 AM"
receivedTime2 = " 06:12 AM"
receivedTimeBoise = " 06:29 AM"
receivedTimeBoise2 = " 06:33 AM"

todaysDateTime = Format(Now(), "ddddd hh:mm AMPM")
x = Len(todaysDateTime)
todaysDate = Left(todaysDateTime, (Len(todaysDateTime) - 9))

'set start and end time based on strings from above'
beginningDate = todaysDate & receivedTime
endingDate = todaysDate & receivedTime2

'clear custorders data'
With Workbooks("Cheese Production Plan.xlsm").Sheets("CustOrders")
    .Range(.Cells(2, 3), .Cells(10000, 28)).Clear
End With

'determine corrrect email'
For Each oOlItm In oOlInb.Items.Restrict("[ReceivedTime] > '" & Format(beginningDate, "ddddd h:nn AMPM") & "' And [ReceivedTime] < '" & Format(endingDate, "ddddd h:nn AMPM") & "'")
    Set oOltargetEmail = oOlItm

        'download attachment to desktop'
        For Each oOlAtch In oOltargetEmail.Attachments
            oOlAtch.SaveAsFile AttachmentPath
        Next

        'open attachment'
        Workbooks.Open (AttachmentPath)
Next

emailFileName = ActiveWorkbook.Name
With Workbooks(emailFileName).Sheets(1)
        .Range("A1:Z1").AutoFilter
        .Range("A1:Z1").AutoFilter Field:=18, Criteria1:="=*CHS*"
        .Range("A1:Z1").AutoFilter Field:=23, Criteria1:="="
        .UsedRange.Offset(1, 0).Resize(.Rows.Count - 1).Copy
End With

With Workbooks("Cheese Production Plan.xlsm").Sheets("CustOrders")
    .Range("C2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End With
Application.CutCopyMode = False
Workbooks(emailFileName).Close SaveChanges:=False

'navigate to boise file
beginningDateBoise = todaysDate & receivedTimeBoise
endingDateBoise = todaysDate & receivedTimeBoise2

For Each oOlItm In oOlInb.Items.Restrict("[ReceivedTime] > '" & Format(beginningDateBoise, "ddddd h:nn AMPM") & "' And [ReceivedTime] < '" & Format(endingDateBoise, "ddddd h:nn AMPM") & "'")
    Set oOltargetEmail = oOlItm

        'download attachment to desktop'
        For Each oOlAtch In oOltargetEmail.Attachments
            oOlAtch.SaveAsFile AttachmentPath2
        Next

        'open attachment'
        Workbooks.Open (AttachmentPath2)
Next
End Sub

Private Sub CommandButton2_Click()
Dim MyObj As New FileSystemObject
Dim MySource As Object
Dim file As Variant
Dim fileDate As String
Dim todayDate As String
Dim t As Integer
Dim firstFile As String
Dim secondFile As String
Dim filePath As String
Dim completePath1 As String
Dim completePath2 As String
Dim targetPath As String
Dim lastInvRow As Integer
Dim x As Integer
Dim futureColumn As Integer
Dim projEndDate As Long
Dim kentFileName As String
Dim boiseWorkbook As String
Dim wkb As Workbook
Dim boiseCount As Long
Dim invPasteMaxRow As Long

For Each wkb In Workbooks
    If Left(wkb.Name, 6) = "boiseE" Then
    boiseWorkbook = wkb.Name
    End If
Next wkb

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = False

Set MySource = MyObj.GetFolder("\\dhqsystemdata\3rd_Party_Warehouse_Prd\WHSE624")
filePath = "\\dhqsystemdata\3rd_Party_Warehouse_Prd\WHSE624\"
t = 0

For Each file In MySource.Files
    fileDate = Int(FileDateTime(file))
    todayDate = Int(Now())

    If (Left(file.Name, 13) = "whse624phyinv") And (fileDate = todayDate) Then
        If t = 0 Then
            firstFile = file.Name
        ElseIf t = 1 Then
            secondFile = file.Name
        End If
        t = t + 1
    Else
    End If
Next file

completePath1 = filePath & firstFile
completePath2 = filePath & secondFile

If FileDateTime(completePath1) > FileDateTime(completePath2) Then
    targetPath = completePath1
Else
    targetPath = completePath2
End If

With Workbooks("Cheese Production Plan.xlsm").Sheets("InventoryPaste")
        .Range(.Cells(2, 8), .Cells(360, 19)).Clear
        .Range(.Cells(3, 1), .Cells(360, 7)).Clear
End With
 'open file, filter and copy'
Workbooks.Open (targetPath)
kentFileName = Application.ActiveWorkbook.Name

With Workbooks(kentFileName).Sheets("Sheet1")
        .Range("A1:K1").AutoFilter
        .Range("A1:K1").AutoFilter Field:=2, Criteria1:="=*CHS*"
        .UsedRange.Offset(1, 0).Resize(.Rows.Count - 1).Copy
End With
'paste in cheese file, fill down formulas and "Kent"'
With Workbooks("Cheese Production Plan.xlsm").Sheets("InventoryPaste")
        .Range("H2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        lastInvRow = .Cells(2, 8).End(xlDown).Row
        'fill Kent'
        .Cells(2, 19).Value = "Kent"
        .Range(.Cells(2, 19), .Cells(2, 19)).AutoFill Destination:=.Range(.Cells(2, 19), .Cells(lastInvRow, 19)), Type:=xlFillDefault
        'fill formulas'
        x = 1
        For x = 1 To 7
                .Range(.Cells(2, x), .Cells(2, x)).AutoFill Destination:=.Range(.Cells(2, x), .Cells(lastInvRow, x)), Type:=xlFillDefault
        Next x
End With
Application.CutCopyMode = False

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

With Workbooks("Cheese Production Plan.xlsm").Sheets("BoiseExtract")
    .Range(.Cells(3, 1), .Cells(250, 22)).Clear
End With

With Workbooks("Cheese Production Plan.xlsm").Sheets("InventoryPaste")
    .Range(.Cells(362, 1), .Cells(500, 19)).Clear
    .Range(.Cells(361, 22), .Cells(361, 32)).Copy
    .Range("H361").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .Range("S361").Value = "Boise"
End With

'do same process to Boise file
With Workbooks(boiseWorkbook).Sheets(1)
        .Range("A1:V1").AutoFilter
        .Range("A1:V1").AutoFilter Field:=5, Criteria1:="=*CHS*"
        .UsedRange.Offset(1, 0).Resize(.Rows.Count - 1).Copy
End With

With Workbooks("Cheese Production Plan.xlsm").Sheets("BoiseExtract")
    .Range("A3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    boiseCount = (.Cells(1, 1).End(xlDown).Row) - 1
End With

invPasteMaxRow = (361 + boiseCount) - 1
x = 1

With Workbooks("Cheese Production Plan.xlsm").Sheets("InventoryPaste")
    For x = 1 To 19
        .Range(.Cells(361, x), .Cells(361, x)).AutoFill Destination:=.Range(.Cells(361, x), .Cells(invPasteMaxRow, x)), Type:=xlFillDefault
    Next x
End With
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Application.CutCopyMode = False
Workbooks(kentFileName).Close SaveChanges:=False
Workbooks("Cheese Production Plan.xlsm").RefreshAll

With Workbooks("Cheese Production Plan.xlsm").Sheets("FutureOrders")
    futureColumn = .Cells(6, 1).End(xlToRight).Column
    projEndDate = .Cells(6, futureColumn - 1).Value
    .Rows("6:6").NumberFormat = "m/d/yyyy"
End With
With Workbooks("Cheese Production Plan.xlsm").Sheets("FGInventory")
    .Cells(9, 21).Value = projEndDate
End With

'/ With Workbooks("Cheese Production Plan.xlsm")
  '/  .Worksheets(Array("FGInventory", "FutureOrders")).Copy
'/ End With

Workbooks(boiseWorkbook).Close SaveChanges:=False
End Sub
