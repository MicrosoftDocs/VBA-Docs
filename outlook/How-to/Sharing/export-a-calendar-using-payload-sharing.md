---
title: Export a Calendar using Payload Sharing
ms.prod: outlook
ms.assetid: acd7d29e-12d6-a5ea-c1a6-8b3165b27dc7
ms.date: 06/08/2017
localization_priority: Normal
---


# Export a Calendar using Payload Sharing

Microsoft Outlook includes the ability to share calendar information with other users by using an iCalendar (.ics) file attached to a **[MailItem](../../../api/Outlook.MailItem.md)**. The **[CalendarSharing](../../../api/Outlook.CalendarSharing.md)** object is used to both generate the iCalendar file from a folder containing calendar items and to generate the **MailItem** to which the iCalendar file is attached.

This sample uses the **CalendarSharing** item to share free/busy information for the next seven days with a single recipient:

1. The sample obtains a **[Folder](../../../api/Outlook.Folder.md)** object reference for the **Calendar** default folder for the current user, by using the **[GetDefaultFolder](../../../api/Outlook.NameSpace.GetDefaultFolder.md)** method of the **[NameSpace](../../../api/Outlook.NameSpace.md)** object.
    
2. It uses the **[GetCalendarExporter](../../../api/Outlook.Folder.GetCalendarExporter.md)** method of the **Folder** object to create a **CalendarSharing** object reference for the folder.
    
3. It then sets the following properties on the **CalendarSharing** object to restrict the range and level of information exported by the object:
    
   - The **[CalendarDetail](../../../api/Outlook.CalendarSharing.CalendarDetail.md)** property is set to limit the information for each calendar item to free/busy information only.
    
   - The **[StartDate](../../../api/Outlook.CalendarSharing.StartDate.md)** and **[EndDate](../../../api/Outlook.CalendarSharing.EndDate.md)** properties are set to restrict the calendar items included in the file to the next seven days.
    
   - The **[RestrictToWorkingHours](../../../api/Outlook.CalendarSharing.RestrictToWorkingHours.md)** property is set to restrict the calendar items to those that fall within working hours.
    
   - The **[IncludeAttachments](../../../api/Outlook.CalendarSharing.IncludeAttachments.md)** property is set to exclude any attachments for calendar items exported by the object.
    
   - The **[IncludePrivateDetails](../../../api/Outlook.CalendarSharing.IncludePrivateDetails.md)** property is set to exclude the details of any private calendar items exported by the object.
    
4. It then calles the **[ForwardAsICal](../../../api/Outlook.CalendarSharing.ForwardAsICal.md)** method of the **CalendarSharing** object to export the calendar items to an iCalendar file and create a **MailItem** object with the iCalendar file as an attachment. The **olCalendarMailFormatDailySchedule** constant of the **[olCalendarMailFormat](../../../api/Outlook.OlCalendarMailFormat.md)** enumeration is used with the **ForwardAsICal** method to indicate that the body of the **MailItem** should contain, in HTML format, free/busy information for the next seven days.
    
5. Finally, the **[Add](../../../api/Outlook.Recipients.Add.md)** method for the **[Recipients](../../../api/Outlook.MailItem.Recipients.md)** collection of the newly created **MailItem** object is called to add the specified recipient and the **[Send](../../../api/Outlook.MailItem.Send(method).md)** method is used to send the **MailItem**.
    

```vb
Public Sub ShareWorkCalendarByPayload() 
 
 Dim oNamespace As NameSpace 
 Dim oFolder As Folder 
 Dim oCalendarSharing As CalendarSharing 
 Dim oMailItem As MailItem 
 
 On Error GoTo ErrRoutine 
 ' Get a reference to the Calendar default folder 
 Set oNamespace = Application.GetNamespace("MAPI") 
 Set oFolder = oNamespace.GetDefaultFolder(olFolderCalendar) 
 
 ' Get a reference to a CalendarSharing object for that 
 ' folder. 
 Set oCalendarSharing = oFolder.GetCalendarExporter 
 
 ' Set the CalendarSharing object to restrict 
 ' the information shared in the iCalendar file. 
 With oCalendarSharing 
 ' Send free/busy information only. 
 .CalendarDetail = olFreeBusyOnly 
 
 ' Send information for the next seven days. 
 .startDate = Now 
 .endDate = DateAdd("d", 7, Now) 
 
 ' Restrict information to working hours only. 
 .RestrictToWorkingHours = True 
 
 ' Exclude attachments and private information. 
 .IncludeAttachments = False 
 .IncludePrivateDetails = False 
 End With 
 
 ' Get the mail item containing the iCalendar file 
 ' and calendar information. 
 Set oMailItem = oCalendarSharing.ForwardAsICal( _ 
 olCalendarMailFormatDailySchedule) 
 
 ' Send the mail item to the specified recipient. 
 With oMailItem 
 .Recipients.Add "someone@example.com" 
 .Send 
 End With 
 
EndRoutine: 
 On Error GoTo 0 
 Set oMailItem = Nothing 
 Set oCalendarSharing = Nothing 
 Set oFolder = Nothing 
 Set oNamespace = Nothing 
Exit Sub 
 
ErrRoutine: 
 Select Case Err.Number 
 Case 287 ' &H0000011F 
 ' The user denied access to the Address Book. 
 ' This error occurs if the code is run by an 
 ' untrusted application, and the user chose not to 
 ' allow access. 
 MsgBox "Access to Outlook was denied by the user.", _ 
 vbOKOnly, _ 
 Err.Number & " - " & Err.Source 
 Case -2147467259 ' &H80004005 
 ' Export failed. 
 ' This error typically occurs if the CalendarSharing 
 ' method cannot export the calendar information because 
 ' of conflicting property settings. 
 MsgBox Err.Description, _ 
 vbOKOnly, _ 
 Err.Number & " - " & Err.Source 
 Case -2147221233 ' &H8004010F 
 ' Operation failed. 
 ' This error typically occurs if the GetCalendarExporter method 
 ' is called on a folder that doesn't contain calendar items. 
 MsgBox Err.Description, _ 
 vbOKOnly, _ 
 Err.Number & " - " & Err.Source 
 Case Else 
 ' Any other error that may occur. 
 MsgBox Err.Description, _ 
 vbOKOnly, _ 
 Err.Number & " - " & Err.Source 
 End Select 
 
 GoTo EndRoutine 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]