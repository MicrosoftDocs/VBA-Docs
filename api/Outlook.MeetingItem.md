---
title: MeetingItem object (Outlook)
keywords: vbaol11.chm2989
f1_keywords:
- vbaol11.chm2989
ms.prod: outlook
api_name:
- Outlook.MeetingItem
ms.assetid: b75730f5-b395-3d66-5acd-b64fd8fcd78f
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem object (Outlook)

Represents a change to the recipient's Calendar folder initiated by another party or as a result of a group action.


## Remarks

Unlike other Microsoft Outlook objects, you cannot create this object. It is created automatically when you set the  **[MeetingStatus](Outlook.AppointmentItem.MeetingStatus.md)** property of an **[AppointmentItem](Outlook.AppointmentItem.md)** object to **olMeeting** and send it to one or more users. They receive it in their inboxes as a **MeetingItem**.

Use the  **[GetAssociatedAppointment](Outlook.MeetingItem.GetAssociatedAppointment.md)** method to return the **AppointmentItem** object associated with a **MeetingItem** object, and work directly with the **AppointmentItem** object to respond to the request.


## Example

The following example uses the  **[CreateItem](Outlook.Application.CreateItem.md)** method to create an appointment. It becomes a **MeetingItem** with both a required and an optional attendee when it is received in the inbox of each of the recipients.


```vb
Set myItem = myOlApp.CreateItem(olAppointmentItem) 
 
myItem.MeetingStatus = olMeeting 
 
myItem.Subject = "Strategy Meeting" 
 
myItem.Location = "Conference Room B" 
 
myItem.Start = #9/24/97 1:30:00 PM# 
 
myItem.Duration = 90 
 
Set myRequiredAttendee = myItem.Recipients.Add("Nate _ 
 
 Sun") 
 
myRequiredAttendee.Type = olRequired 
 
Set myOptionalAttendee = myItem.Recipients.Add("Kevin _ 
 
 Kennedy") 
 
myOptionalAttendee.Type = olOptional 
 
Set myResourceAttendee = _ 
 
 myItem.Recipients.Add("Conference Room B") 
 
myResourceAttendee.Type = olResource 
 
myItem.Send
```


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.MeetingItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.MeetingItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.MeetingItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.MeetingItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.MeetingItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.MeetingItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.MeetingItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.MeetingItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.MeetingItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.MeetingItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.MeetingItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.MeetingItem.BeforeDelete.md)|
|[BeforeRead](Outlook.MeetingItem.BeforeRead.md)|
|[Close](Outlook.MeetingItem.Close(even).md)|
|[CustomAction](Outlook.MeetingItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.MeetingItem.CustomPropertyChange.md)|
|[Forward](Outlook.MeetingItem.Forward(even).md)|
|[Open](Outlook.MeetingItem.Open.md)|
|[PropertyChange](Outlook.MeetingItem.PropertyChange.md)|
|[Read](Outlook.MeetingItem.Read.md)|
|[ReadComplete](Outlook.meetingitem.readcomplete.md)|
|[Reply](Outlook.MeetingItem.Reply(even).md)|
|[ReplyAll](Outlook.MeetingItem.ReplyAll(even).md)|
|[Send](Outlook.MeetingItem.Send(even).md)|
|[Unload](Outlook.MeetingItem.Unload.md)|
|[Write](Outlook.MeetingItem.Write.md)|

## Methods



|Name|
|:-----|
|[Close](Outlook.MeetingItem.Close(method).md)|
|[Copy](Outlook.MeetingItem.Copy.md)|
|[Delete](Outlook.MeetingItem.Delete.md)|
|[Display](Outlook.MeetingItem.Display.md)|
|[Forward](Outlook.MeetingItem.Forward(method).md)|
|[GetAssociatedAppointment](Outlook.MeetingItem.GetAssociatedAppointment.md)|
|[GetConversation](Outlook.MeetingItem.GetConversation.md)|
|[Move](Outlook.MeetingItem.Move.md)|
|[PrintOut](Outlook.MeetingItem.PrintOut.md)|
|[Reply](Outlook.MeetingItem.Reply(method).md)|
|[ReplyAll](Outlook.MeetingItem.ReplyAll(method).md)|
|[Save](Outlook.MeetingItem.Save.md)|
|[SaveAs](Outlook.MeetingItem.SaveAs.md)|
|[Send](Outlook.MeetingItem.Send(method).md)|
|[ShowCategoriesDialog](Outlook.MeetingItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.MeetingItem.Actions.md)|
|[Application](Outlook.MeetingItem.Application.md)|
|[Attachments](Outlook.MeetingItem.Attachments.md)|
|[AutoForwarded](Outlook.MeetingItem.AutoForwarded.md)|
|[AutoResolvedWinner](Outlook.MeetingItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.MeetingItem.BillingInformation.md)|
|[Body](Outlook.MeetingItem.Body.md)|
|[Categories](Outlook.MeetingItem.Categories.md)|
|[Class](Outlook.MeetingItem.Class.md)|
|[Companies](Outlook.MeetingItem.Companies.md)|
|[Conflicts](Outlook.MeetingItem.Conflicts.md)|
|[ConversationID](Outlook.MeetingItem.ConversationID.md)|
|[ConversationIndex](Outlook.MeetingItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.MeetingItem.ConversationTopic.md)|
|[CreationTime](Outlook.MeetingItem.CreationTime.md)|
|[DeferredDeliveryTime](Outlook.MeetingItem.DeferredDeliveryTime.md)|
|[DeleteAfterSubmit](Outlook.MeetingItem.DeleteAfterSubmit.md)|
|[DownloadState](Outlook.MeetingItem.DownloadState.md)|
|[EntryID](Outlook.MeetingItem.EntryID.md)|
|[ExpiryTime](Outlook.MeetingItem.ExpiryTime.md)|
|[FormDescription](Outlook.MeetingItem.FormDescription.md)|
|[GetInspector](Outlook.MeetingItem.GetInspector.md)|
|[Importance](Outlook.MeetingItem.Importance.md)|
|[IsConflict](Outlook.MeetingItem.IsConflict.md)|
|[IsLatestVersion](Outlook.MeetingItem.IsLatestVersion.md)|
|[ItemProperties](Outlook.MeetingItem.ItemProperties.md)|
|[LastModificationTime](Outlook.MeetingItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.MeetingItem.MarkForDownload.md)|
|[MeetingWorkspaceURL](Outlook.MeetingItem.MeetingWorkspaceURL.md)|
|[MessageClass](Outlook.MeetingItem.MessageClass.md)|
|[Mileage](Outlook.MeetingItem.Mileage.md)|
|[NoAging](Outlook.MeetingItem.NoAging.md)|
|[OriginatorDeliveryReportRequested](Outlook.MeetingItem.OriginatorDeliveryReportRequested.md)|
|[OutlookInternalVersion](Outlook.MeetingItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.MeetingItem.OutlookVersion.md)|
|[Parent](Outlook.MeetingItem.Parent.md)|
|[PropertyAccessor](Outlook.MeetingItem.PropertyAccessor.md)|
|[ReceivedTime](Outlook.MeetingItem.ReceivedTime.md)|
|[Recipients](Outlook.MeetingItem.Recipients.md)|
|[ReminderSet](Outlook.MeetingItem.ReminderSet.md)|
|[ReminderTime](Outlook.MeetingItem.ReminderTime.md)|
|[ReplyRecipients](Outlook.MeetingItem.ReplyRecipients.md)|
|[RetentionExpirationDate](Outlook.MeetingItem.RetentionExpirationDate.md)|
|[RetentionPolicyName](Outlook.MeetingItem.RetentionPolicyName.md)|
|[RTFBody](Outlook.MeetingItem.RTFBody.md)|
|[Saved](Outlook.MeetingItem.Saved.md)|
|[SaveSentMessageFolder](Outlook.MeetingItem.SaveSentMessageFolder.md)|
|[SenderEmailAddress](Outlook.MeetingItem.SenderEmailAddress.md)|
|[SenderEmailType](Outlook.MeetingItem.SenderEmailType.md)|
|[SenderName](Outlook.MeetingItem.SenderName.md)|
|[SendUsingAccount](Outlook.MeetingItem.SendUsingAccount.md)|
|[Sensitivity](Outlook.MeetingItem.Sensitivity.md)|
|[Sent](Outlook.MeetingItem.Sent.md)|
|[SentOn](Outlook.MeetingItem.SentOn.md)|
|[Session](Outlook.MeetingItem.Session.md)|
|[Size](Outlook.MeetingItem.Size.md)|
|[Subject](Outlook.MeetingItem.Subject.md)|
|[Submitted](Outlook.MeetingItem.Submitted.md)|
|[UnRead](Outlook.MeetingItem.UnRead.md)|
|[UserProperties](Outlook.MeetingItem.UserProperties.md)|

## See also


[MeetingItem Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
