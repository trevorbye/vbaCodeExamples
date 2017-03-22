Const olFolderInbox As Integer = 6

Application.ScreenUpdating = False
Dim oOlAp As Object, oOlns As Object, oOlInb As Object, oOlItm As Object, oOltargetEmail As Object, oOlAtch As Object
Dim beginningDate As String, endingDate As String, receivedTime As String, receivedTime2 As String, start As String, endTime As String
Dim counter As Integer


Set oOlAp = GetObject(, "Outlook.application")
Set oOlns = oOlAp.GetNamespace("MAPI")
Set oOlInb = oOlns.GetDefaultFolder(olFolderInbox)

start = "3/14/2017"
receivedTime = " 06:00 AM"

endTime = "3/15/2017"
receivedTime2 = " 08:00 AM"

beginningDate = start & receivedTime
endingDate = endTime & receivedTime2

counter = 0
For Each oOlItm In oOlInb.Items.Restrict("[ReceivedTime] > '" & Format(beginningDate, "ddddd h:nn AMPM") & "' And [ReceivedTime] < '" & Format(endingDate, "ddddd h:nn AMPM") & "'")

	Dim body As String
	Dim sender As String
	Dim sentDate As String
	Dim subject As String
	Dim recipients As String
    Set oOltargetEmail = oOlItm

    body = oOltargetEmail.Body
    sender = oOltargetEmail.SenderEmailAddress
    sentDate = oOltargetEmail.SentOn
    subject = oOltargetEmail.Subject
    recipients = oOltargetEmail.To

    If sender = "appodgp@darigold.com" Then
    	GoTo nullProcess
    Else
    	'do stuff'
    End If

nullProcess: 
counter = counter + 1  

'time control'
If counter > 50 Then
	Exit For
End If 

Next