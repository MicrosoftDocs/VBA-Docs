---
title: AppointmentItem object (Outlook)
keywords: vbaol11.chm2988
f1_keywords:
- vbaol11.chm2988
ms.prod: outlook
api_name:
- Outlook.AppointmentItem
ms.assetid: 204a409d-654e-27aa-643a-8344c631b82d
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem object (Outlook)

Represents a meeting, a one-time appointment, or a recurring appointment or meeting in the Calendar folder.


## Remarks

Use the  **[CreateItem](Outlook.Application.CreateItem.md)** method to create an **AppointmentItem** object that represents a new appointment.

Use  **[Items](Outlook.Items.Item.md)** (_index_), where _index_ is the index number of an appointment or a value used to match the default property of an appointment, to return a single **AppointmentItem** object from a Calendar folder.

You can also return an **AppointmentItem** object from a **[MeetingItem](Outlook.MeetingItem.md)** object by using the **[GetAssociatedAppointment](Outlook.MeetingItem.GetAssociatedAppointment.md)** method.

When you work with recurring appointment items, you should release any prior references, obtain new references to the recurring appointment item before you access or modify the item, and release these references as soon as you are finished and have saved the changes. This practice applies to the recurring  **AppointmentItem** object, and any **[Exception](Outlook.Exception.md)** or **[RecurrencePattern](Outlook.RecurrencePattern.md)** object. To release a reference in Visual Basic for Applications (VBA) or Visual Basic, set that existing object to **Nothing**. In C#, explicitly release the memory for that object.

Note that even after you release your reference and attempt to obtain a new reference, if there is still an active reference, held by another add-in or Outlook, to one of the above objects, your new reference will still point to an out-of-date copy of the object. Therefore, it is important that you release your references as soon as you are finished with the recurring appointment.

The following code example in VBA shows how to release and refresh references in order to obtain up-to-date data for a recurring appointment. The example obtains a set of appointment items from the Calendar folder. It assumes that the first item in the appointment collection is part of a recurring appointment. The example shows that a reference to the appointment collection obtained before an exception is created does not reflect the exception. The example then releases this reference and other existing appointment references, after which new references that point to the appointment collection reflect the exception.




```vb
Sub TestExceptions() 
 
 Dim oItems As Items 
 
 Dim oItemOriginal As AppointmentItem 
 
 Dim oItemNew As AppointmentItem 
 
 Dim rPattern As RecurrencePattern 
 
 Dim oEx As Exceptions 
 
 Dim oEx2 As Exceptions 
 
 Dim oOccurrence As AppointmentItem 
 
 Dim i As Long 
 
 
 
 ' This is the initial reference to an appointment collection. 
 
 Set oItems = _ 
 
 Outlook.Application.Session.GetDefaultFolder(olFolderCalendar).Items 
 
 
 
 ' This is the original reference to the first appointment in the 
 
 ' collection before an exception is created. 
 
 Set oItemOriginal = oItems.Item(1) 
 
 
 
 ' Code example assumes that the first appointment in the collection 
 
 ' is a recurring appointment. 
 
 Set oOccurrence = _ 
 
 oItemOriginal.GetRecurrencePattern().GetOccurrence(#2/28/2010 8:00:00 AM#) 
 
 
 
 ' Create an exception by changing the 2/28 occurrence to 3/3. 
 
 oOccurrence.Start = #3/3/2010 8:00:00 AM# 
 
 oOccurrence.Save 
 
 
 
 Stop 
 
 
 
 ' Preexisting reference to the first appointment in the collection 
 
 ' does not reflect the exception. 
 
 oItemOriginal.Save 
 
 Set oEx = oItemOriginal.GetRecurrencePattern().Exceptions 
 
 Debug.Print oItemOriginal.subject 
 
 Debug.Print " Original item exceptions: " & oEx.Count 
 
 
 
 ' Get a new reference based on the existing reference to the 
 
 ' appointment collection created before the exception. 
 
 ' The new reference does not reflect the exception. 
 
 Set oItemNew = oItems.Item(1) 
 
 oItemNew.Save 
 
 Set oEx2 = oItemNew.GetRecurrencePattern().Exceptions 
 
 Debug.Print " New item exceptions: " & oEx2.Count 
 
 
 
 ' Same: preexisting reference to the first appointment in the collection 
 
 ' does not reflect the exception. 
 
 Set oEx = oItemOriginal.GetRecurrencePattern().Exceptions 
 
 Debug.Print " Original item exceptions: " & oEx.Count 
 
 
 
 ' Release all existing references to appointment items, 
 
 ' including the appointment collection, an exception, occurrence, 
 
 ' or any other appointment. 
 
 Debug.Print "REFRESH ITEM COLLECTION" 
 
 Set oItems = Nothing 
 
 Set oItemNew = Nothing 
 
 Set oEx = Nothing 
 
 Set oEx2 = Nothing 
 
 Set oOccurrence = Nothing 
 
 Set oItemOriginal = Nothing 
 
 Set rPattern = Nothing 
 
 
 
 ' Get new references to appointment items, including the appointment 
 
 ' collection, individual appointments, and exceptions. 
 
 Set oItems = _ 
 
 Outlook.Application.Session.GetDefaultFolder(olFolderCalendar).Items 
 
 Set oItemNew = oItems.Item(1) 
 
 
 
 ' If no other add-ins have the same recurring appointment open, 
 
 ' the new references reflect the current exception count. 
 
 Set oEx2 = oItemNew.GetRecurrencePattern().Exceptions 
 
 Debug.Print " New item exceptions: " & oEx2.Count 
 
 
 
 Debug.Print "RE-GET ORIGINAL" 
 
 Set oItemOriginal = oItems.Item(1) 
 
 Set oEx = oItemOriginal.GetRecurrencePattern().Exceptions 
 
 Debug.Print " Original item exceptions: " & oEx.Count 
 
End Sub
```


## Example

The following Visual Basic for Applications (VBA) example returns a new appointment.


```vb
Set myItem = Application.CreateItem(olAppointmentItem)
```


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.AppointmentItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.AppointmentItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.AppointmentItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.AppointmentItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.AppointmentItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.AppointmentItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.AppointmentItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.AppointmentItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.AppointmentItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.AppointmentItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.AppointmentItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.AppointmentItem.BeforeDelete.md)|
|[BeforeRead](Outlook.AppointmentItem.BeforeRead.md)|
|[Close](Outlook.AppointmentItem.Close(even).md)|
|[CustomAction](Outlook.AppointmentItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.AppointmentItem.CustomPropertyChange.md)|
|[Forward](Outlook.AppointmentItem.Forward.md)|
|[Open](Outlook.AppointmentItem.Open.md)|
|[PropertyChange](Outlook.AppointmentItem.PropertyChange.md)|
|[Read](Outlook.AppointmentItem.Read.md)|
|[ReadComplete](Outlook.appointmentitem.readcomplete.md)|
|[Reply](Outlook.AppointmentItem.Reply.md)|
|[ReplyAll](Outlook.AppointmentItem.ReplyAll.md)|
|[Send](Outlook.AppointmentItem.Send(even).md)|
|[Unload](Outlook.AppointmentItem.Unload.md)|
|[Write](Outlook.AppointmentItem.Write.md)|

## Methods



|Name|
|:-----|
|[ClearRecurrencePattern](Outlook.AppointmentItem.ClearRecurrencePattern.md)|
|[Close](Outlook.AppointmentItem.Close(method).md)|
|[Copy](Outlook.AppointmentItem.Copy.md)|
|[CopyTo](Outlook.AppointmentItem.CopyTo.md)|
|[Delete](Outlook.AppointmentItem.Delete.md)|
|[Display](Outlook.AppointmentItem.Display.md)|
|[ForwardAsVcal](Outlook.AppointmentItem.ForwardAsVcal.md)|
|[GetConversation](Outlook.AppointmentItem.GetConversation.md)|
|[GetOrganizer](Outlook.AppointmentItem.GetOrganizer.md)|
|[GetRecurrencePattern](Outlook.AppointmentItem.GetRecurrencePattern.md)|
|[Move](Outlook.AppointmentItem.Move.md)|
|[PrintOut](Outlook.AppointmentItem.PrintOut.md)|
|[Respond](Outlook.AppointmentItem.Respond.md)|
|[Save](Outlook.AppointmentItem.Save.md)|
|[SaveAs](Outlook.AppointmentItem.SaveAs.md)|
|[Send](Outlook.AppointmentItem.Send(method).md)|
|[ShowCategoriesDialog](Outlook.AppointmentItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.AppointmentItem.Actions.md)|
|[AllDayEvent](Outlook.AppointmentItem.AllDayEvent.md)|
|[Application](Outlook.AppointmentItem.Application.md)|
|[Attachments](Outlook.AppointmentItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.AppointmentItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.AppointmentItem.BillingInformation.md)|
|[Body](Outlook.AppointmentItem.Body.md)|
|[BusyStatus](Outlook.AppointmentItem.BusyStatus.md)|
|[Categories](Outlook.AppointmentItem.Categories.md)|
|[Class](Outlook.AppointmentItem.Class.md)|
|[Companies](Outlook.AppointmentItem.Companies.md)|
|[Conflicts](Outlook.AppointmentItem.Conflicts.md)|
|[ConversationID](Outlook.AppointmentItem.ConversationID.md)|
|[ConversationIndex](Outlook.AppointmentItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)|
|[CreationTime](Outlook.AppointmentItem.CreationTime.md)|
|[DownloadState](Outlook.AppointmentItem.DownloadState.md)|
|[Duration](Outlook.AppointmentItem.Duration.md)|
|[End](Outlook.AppointmentItem.End.md)|
|[EndInEndTimeZone](Outlook.AppointmentItem.EndInEndTimeZone.md)|
|[EndTimeZone](Outlook.AppointmentItem.EndTimeZone.md)|
|[EndUTC](Outlook.AppointmentItem.EndUTC.md)|
|[EntryID](Outlook.AppointmentItem.EntryID.md)|
|[ForceUpdateToAllAttendees](Outlook.AppointmentItem.ForceUpdateToAllAttendees.md)|
|[FormDescription](Outlook.AppointmentItem.FormDescription.md)|
|[GetInspector](Outlook.AppointmentItem.GetInspector.md)|
|[GlobalAppointmentID](Outlook.AppointmentItem.GlobalAppointmentID.md)|
|[Importance](Outlook.AppointmentItem.Importance.md)|
|[InternetCodepage](Outlook.AppointmentItem.InternetCodepage.md)|
|[IsConflict](Outlook.AppointmentItem.IsConflict.md)|
|[IsRecurring](Outlook.AppointmentItem.IsRecurring.md)|
|[ItemProperties](Outlook.AppointmentItem.ItemProperties.md)|
|[LastModificationTime](Outlook.AppointmentItem.LastModificationTime.md)|
|[Location](Outlook.AppointmentItem.Location.md)|
|[MarkForDownload](Outlook.AppointmentItem.MarkForDownload.md)|
|[MeetingStatus](Outlook.AppointmentItem.MeetingStatus.md)|
|[MeetingWorkspaceURL](Outlook.AppointmentItem.MeetingWorkspaceURL.md)|
|[MessageClass](Outlook.AppointmentItem.MessageClass.md)|
|[Mileage](Outlook.AppointmentItem.Mileage.md)|
|[NoAging](Outlook.AppointmentItem.NoAging.md)|
|[OptionalAttendees](Outlook.AppointmentItem.OptionalAttendees.md)|
|[Organizer](Outlook.AppointmentItem.Organizer.md)|
|[OutlookInternalVersion](Outlook.AppointmentItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.AppointmentItem.OutlookVersion.md)|
|[Parent](Outlook.AppointmentItem.Parent.md)|
|[PropertyAccessor](Outlook.AppointmentItem.PropertyAccessor.md)|
|[Recipients](Outlook.AppointmentItem.Recipients.md)|
|[RecurrenceState](Outlook.AppointmentItem.RecurrenceState.md)|
|[ReminderMinutesBeforeStart](Outlook.AppointmentItem.ReminderMinutesBeforeStart.md)|
|[ReminderOverrideDefault](Outlook.AppointmentItem.ReminderOverrideDefault.md)|
|[ReminderPlaySound](Outlook.AppointmentItem.ReminderPlaySound.md)|
|[ReminderSet](Outlook.AppointmentItem.ReminderSet.md)|
|[ReminderSoundFile](Outlook.AppointmentItem.ReminderSoundFile.md)|
|[ReplyTime](Outlook.AppointmentItem.ReplyTime.md)|
|[RequiredAttendees](Outlook.AppointmentItem.RequiredAttendees.md)|
|[Resources](Outlook.AppointmentItem.Resources.md)|
|[ResponseRequested](Outlook.AppointmentItem.ResponseRequested.md)|
|[ResponseStatus](Outlook.AppointmentItem.ResponseStatus.md)|
|[RTFBody](Outlook.AppointmentItem.RTFBody.md)|
|[Saved](Outlook.AppointmentItem.Saved.md)|
|[SendUsingAccount](Outlook.AppointmentItem.SendUsingAccount.md)|
|[Sensitivity](Outlook.AppointmentItem.Sensitivity.md)|
|[Session](Outlook.AppointmentItem.Session.md)|
|[Size](Outlook.AppointmentItem.Size.md)|
|[Start](Outlook.AppointmentItem.Start.md)|
|[StartInStartTimeZone](Outlook.AppointmentItem.StartInStartTimeZone.md)|
|[StartTimeZone](Outlook.AppointmentItem.StartTimeZone.md)|
|[StartUTC](Outlook.AppointmentItem.StartUTC.md)|
|[Subject](Outlook.AppointmentItem.Subject.md)|
|[UnRead](Outlook.AppointmentItem.UnRead.md)|
|[UserProperties](Outlook.AppointmentItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[How to: Import Appointment XML Data into Outlook Appointment Objects](../outlook/How-to/Items-Folders-and-Stores/import-appointment-xml-data-into-outlook-appointment-objects-outlook.md)
[AppointmentItem Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
