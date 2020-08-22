Option Explicit
Public Sub CreateOutlookAppointments() 'from https://www.slipstick.com/developer/create-appointments-spreadsheet-data/

   Sheets("Sheet1").Select
    On Error GoTo Err_Execute
     
    Dim olApp As Outlook.Application
    Dim olAppt As Outlook.AppointmentItem
    Dim blnCreated As Boolean
    Dim olNs As Outlook.Namespace
    Dim CalFolder As Outlook.MAPIFolder
     
    Dim i As Long
     
    On Error Resume Next
    Set olApp = Outlook.Application
     
    If olApp Is Nothing Then
        Set olApp = Outlook.Application
         blnCreated = True
        Err.Clear
    Else
        blnCreated = False
    End If
     
    On Error GoTo 0
     
    Set olNs = olApp.GetNamespace("MAPI")
    Set CalFolder = olNs.GetDefaultFolder(olFolderCalendar)
         
    i = 2
    Do Until Trim(Cells(i, 1).Value) = ""
    
    Dim EmailString As String 'ADDED BY DANI!!!!!!!!!!!!!!!!!!!!!!!!
    EmailString = Cells(i, 9) 'ADDED BY DANI!!!!!!!!!!!!!!!!!!!!!!!!
    Dim EmailArray() As String 'ADDED BY DANI!!!!!!!!!!!!!!!!!!!!!!!!
    EmailArray = Split(EmailString, ";") 'ADDED BY DANI!!!!!!!!!!!!!!!!!!!!!!!!
    
    Set olAppt = CalFolder.Items.Add(olAppointmentItem)
    
    With olAppt
        .MeetingStatus = olMeeting
        
    'Define calendar item properties
        .Start = Cells(i, 4) + Cells(i, 5)
        .End = Cells(i, 6) + Cells(i, 7)
        .Subject = Cells(i, 1)
        .Location = Cells(i, 2)
        .Body = Cells(i, 3)
        .BusyStatus = olBusy
        .ReminderMinutesBeforeStart = Cells(i, 8)
        .ReminderSet = True
        Dim RequiredAttendee As Outlook.Recipient
            
            Dim j As Integer
            j = 0 'ADDED BY DANI!!!!!!!!!!!!!!!!!!!!!!!!
            Dim Length As Integer
            Length = UBound(EmailArray) - LBound(EmailArray) + 1
            Do Until j = Length 'ADDED BY DANI!!!!!!!!!!!!!!!!!!!!!!!!
            
            Set RequiredAttendee = .Recipients.Add(EmailArray(j))
                RequiredAttendee.Type = olRequired
            'need to add a split function and loop through the attendee code if we want to remind more than one person
            j = j + 1 'ADDED BY DANI!!!!!!!!!!!!!!!!!!!!!!!!
            Loop 'ADDED BY DANI!!!!!!!!!!!!!!!!!!!!!!!!
        
        .Display
    End With
                 
        i = i + 1
        Loop
    Set olAppt = Nothing
    Set olApp = Nothing
     
    Exit Sub
     
Err_Execute:
    MsgBox "An error occurred - Exporting items to Calendar."
     
End Sub

