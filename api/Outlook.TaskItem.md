---
title: TaskItem object (Outlook)
keywords: vbaol11.chm2990
f1_keywords:
- vbaol11.chm2990
ms.prod: outlook
api_name:
- Outlook.TaskItem
ms.assetid: 5df8cfa5-5460-a5a1-a130-ba5bca1a0091
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem object (Outlook)

Represents a task (an assigned, delegated, or self-imposed task to be performed within a specified time frame) in a Tasks folder.


## Remarks

Use the  **[CreateItem](Outlook.Application.CreateItem.md)** method to create a **TaskItem** object that represents a new task.

Use  **[Items](Outlook.Folder.Items.md)** (_index_), where _index_ is the index number of a task or a value used to match the default property of a task, to return a single **TaskItem** object from a Tasks folder.


## Example

The following Visual Basic for Applications (VBA) example returns a new task.






```vb
Set myItem = Application.CreateItem(olTaskItem)
```


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.TaskItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.TaskItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.TaskItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.TaskItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.TaskItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.TaskItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.TaskItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.TaskItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.TaskItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.TaskItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.TaskItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.TaskItem.BeforeDelete.md)|
|[BeforeRead](Outlook.TaskItem.BeforeRead.md)|
|[Close](Outlook.TaskItem.Close(even).md)|
|[CustomAction](Outlook.TaskItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.TaskItem.CustomPropertyChange.md)|
|[Forward](Outlook.TaskItem.Forward.md)|
|[Open](Outlook.TaskItem.Open.md)|
|[PropertyChange](Outlook.TaskItem.PropertyChange.md)|
|[Read](Outlook.TaskItem.Read.md)|
|[ReadComplete](Outlook.taskitem.readcomplete.md)|
|[Reply](Outlook.TaskItem.Reply.md)|
|[ReplyAll](Outlook.TaskItem.ReplyAll.md)|
|[Send](Outlook.TaskItem.Send(even).md)|
|[Unload](Outlook.TaskItem.Unload.md)|
|[Write](Outlook.TaskItem.Write.md)|

## Methods



|Name|
|:-----|
|[Assign](Outlook.TaskItem.Assign.md)|
|[CancelResponseState](Outlook.TaskItem.CancelResponseState.md)|
|[ClearRecurrencePattern](Outlook.TaskItem.ClearRecurrencePattern.md)|
|[Close](Outlook.TaskItem.Close(method).md)|
|[Copy](Outlook.TaskItem.Copy.md)|
|[Delete](Outlook.TaskItem.Delete.md)|
|[Display](Outlook.TaskItem.Display.md)|
|[GetConversation](Outlook.TaskItem.GetConversation.md)|
|[GetRecurrencePattern](Outlook.TaskItem.GetRecurrencePattern.md)|
|[MarkComplete](Outlook.TaskItem.MarkComplete.md)|
|[Move](Outlook.TaskItem.Move.md)|
|[PrintOut](Outlook.TaskItem.PrintOut.md)|
|[Respond](Outlook.TaskItem.Respond.md)|
|[Save](Outlook.TaskItem.Save.md)|
|[SaveAs](Outlook.TaskItem.SaveAs.md)|
|[Send](Outlook.TaskItem.Send(method).md)|
|[ShowCategoriesDialog](Outlook.TaskItem.ShowCategoriesDialog.md)|
|[SkipRecurrence](Outlook.TaskItem.SkipRecurrence.md)|
|[StatusReport](Outlook.TaskItem.StatusReport.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.TaskItem.Actions.md)|
|[ActualWork](Outlook.TaskItem.ActualWork.md)|
|[Application](Outlook.TaskItem.Application.md)|
|[Attachments](Outlook.TaskItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.TaskItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.TaskItem.BillingInformation.md)|
|[Body](Outlook.TaskItem.Body.md)|
|[CardData](Outlook.TaskItem.CardData.md)|
|[Categories](Outlook.TaskItem.Categories.md)|
|[Class](Outlook.TaskItem.Class.md)|
|[Companies](Outlook.TaskItem.Companies.md)|
|[Complete](Outlook.TaskItem.Complete.md)|
|[Conflicts](Outlook.TaskItem.Conflicts.md)|
|[ContactNames](Outlook.TaskItem.ContactNames.md)|
|[ConversationID](Outlook.TaskItem.ConversationID.md)|
|[ConversationIndex](Outlook.TaskItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.TaskItem.ConversationTopic.md)|
|[CreationTime](Outlook.TaskItem.CreationTime.md)|
|[DateCompleted](Outlook.TaskItem.DateCompleted.md)|
|[DelegationState](Outlook.TaskItem.DelegationState.md)|
|[Delegator](Outlook.TaskItem.Delegator.md)|
|[DownloadState](Outlook.TaskItem.DownloadState.md)|
|[DueDate](Outlook.TaskItem.DueDate.md)|
|[EntryID](Outlook.TaskItem.EntryID.md)|
|[FormDescription](Outlook.TaskItem.FormDescription.md)|
|[GetInspector](Outlook.TaskItem.GetInspector.md)|
|[Importance](Outlook.TaskItem.Importance.md)|
|[InternetCodepage](Outlook.TaskItem.InternetCodepage.md)|
|[IsConflict](Outlook.TaskItem.IsConflict.md)|
|[IsRecurring](Outlook.TaskItem.IsRecurring.md)|
|[ItemProperties](Outlook.TaskItem.ItemProperties.md)|
|[LastModificationTime](Outlook.TaskItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.TaskItem.MarkForDownload.md)|
|[MessageClass](Outlook.TaskItem.MessageClass.md)|
|[Mileage](Outlook.TaskItem.Mileage.md)|
|[NoAging](Outlook.TaskItem.NoAging.md)|
|[Ordinal](Outlook.TaskItem.Ordinal.md)|
|[OutlookInternalVersion](Outlook.TaskItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.TaskItem.OutlookVersion.md)|
|[Owner](Outlook.TaskItem.Owner.md)|
|[Ownership](Outlook.TaskItem.Ownership.md)|
|[Parent](Outlook.TaskItem.Parent.md)|
|[PercentComplete](Outlook.TaskItem.PercentComplete.md)|
|[PropertyAccessor](Outlook.TaskItem.PropertyAccessor.md)|
|[Recipients](Outlook.TaskItem.Recipients.md)|
|[ReminderOverrideDefault](Outlook.TaskItem.ReminderOverrideDefault.md)|
|[ReminderPlaySound](Outlook.TaskItem.ReminderPlaySound.md)|
|[ReminderSet](Outlook.TaskItem.ReminderSet.md)|
|[ReminderSoundFile](Outlook.TaskItem.ReminderSoundFile.md)|
|[ReminderTime](Outlook.TaskItem.ReminderTime.md)|
|[ResponseState](Outlook.TaskItem.ResponseState.md)|
|[Role](Outlook.TaskItem.Role.md)|
|[RTFBody](Outlook.TaskItem.RTFBody.md)|
|[Saved](Outlook.TaskItem.Saved.md)|
|[SchedulePlusPriority](Outlook.TaskItem.SchedulePlusPriority.md)|
|[SendUsingAccount](Outlook.TaskItem.SendUsingAccount.md)|
|[Sensitivity](Outlook.TaskItem.Sensitivity.md)|
|[Session](Outlook.TaskItem.Session.md)|
|[Size](Outlook.TaskItem.Size.md)|
|[StartDate](Outlook.TaskItem.StartDate.md)|
|[Status](Outlook.TaskItem.Status.md)|
|[StatusOnCompletionRecipients](Outlook.TaskItem.StatusOnCompletionRecipients.md)|
|[StatusUpdateRecipients](Outlook.TaskItem.StatusUpdateRecipients.md)|
|[Subject](Outlook.TaskItem.Subject.md)|
|[TeamTask](Outlook.TaskItem.TeamTask.md)|
|[ToDoTaskOrdinal](Outlook.TaskItem.ToDoTaskOrdinal.md)|
|[TotalWork](Outlook.TaskItem.TotalWork.md)|
|[UnRead](Outlook.TaskItem.UnRead.md)|
|[UserProperties](Outlook.TaskItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[TaskItem Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
