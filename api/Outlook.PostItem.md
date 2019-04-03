---
title: PostItem object (Outlook)
keywords: vbaol11.chm3005
f1_keywords:
- vbaol11.chm3005
ms.prod: outlook
api_name:
- Outlook.PostItem
ms.assetid: de44065d-4e93-315a-279f-7b92f09c0465
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem object (Outlook)

Represents a post in a public folder that others may browse.


## Remarks

Unlike a  **[MailItem](Outlook.MailItem.md)** object, a **PostItem** object is not sent to a recipient. You use the **[Post](Outlook.PostItem.Post.md)** method, which is analogous to the **[Send](Outlook.MailItem.Send(method).md)** method for the **MailItem** object, to save the **PostItem** to the target public folder instead of mailing it.

Use the  **[CreateItem](Outlook.Application.CreateItem.md)** or **[CreateItemFromTemplate](Outlook.Application.CreateItemFromTemplate.md)** method to create a **PostItem** object that represents a new post.

Use  **[Items](Outlook.Items.md)** (_index_), where _index_ is the index number of a post or a value used to match the default property of a post, to return a single **PostItem** object from a public folder.


## Example

The following example returns a new post.


```vb
Set myItem = myOlApp.CreateItem(olPostItem)
```


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.PostItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.PostItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.PostItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.PostItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.PostItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.PostItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.PostItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.PostItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.PostItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.PostItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.PostItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.PostItem.BeforeDelete.md)|
|[BeforeRead](Outlook.PostItem.BeforeRead.md)|
|[Close](Outlook.PostItem.Close(even).md)|
|[CustomAction](Outlook.PostItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.PostItem.CustomPropertyChange.md)|
|[Forward](Outlook.PostItem.Forward(even).md)|
|[Open](Outlook.PostItem.Open.md)|
|[PropertyChange](Outlook.PostItem.PropertyChange.md)|
|[Read](Outlook.PostItem.Read.md)|
|[ReadComplete](Outlook.postitem.readcomplete.md)|
|[Reply](Outlook.PostItem.Reply(even).md)|
|[ReplyAll](Outlook.PostItem.ReplyAll.md)|
|[Send](Outlook.PostItem.Send.md)|
|[Unload](Outlook.PostItem.Unload.md)|
|[Write](Outlook.PostItem.Write.md)|

## Methods



|Name|
|:-----|
|[ClearConversationIndex](Outlook.PostItem.ClearConversationIndex.md)|
|[ClearTaskFlag](Outlook.PostItem.ClearTaskFlag.md)|
|[Close](Outlook.PostItem.Close(method).md)|
|[Copy](Outlook.PostItem.Copy.md)|
|[Delete](Outlook.PostItem.Delete.md)|
|[Display](Outlook.PostItem.Display.md)|
|[Forward](Outlook.PostItem.Forward(method).md)|
|[GetConversation](Outlook.PostItem.GetConversation.md)|
|[MarkAsTask](Outlook.PostItem.MarkAsTask.md)|
|[Move](Outlook.PostItem.Move.md)|
|[Post](Outlook.PostItem.Post.md)|
|[PrintOut](Outlook.PostItem.PrintOut.md)|
|[Reply](Outlook.PostItem.Reply(method).md)|
|[Save](Outlook.PostItem.Save.md)|
|[SaveAs](Outlook.PostItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.PostItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.PostItem.Actions.md)|
|[Application](Outlook.PostItem.Application.md)|
|[Attachments](Outlook.PostItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.PostItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.PostItem.BillingInformation.md)|
|[Body](Outlook.PostItem.Body.md)|
|[BodyFormat](Outlook.PostItem.BodyFormat.md)|
|[Categories](Outlook.PostItem.Categories.md)|
|[Class](Outlook.PostItem.Class.md)|
|[Companies](Outlook.PostItem.Companies.md)|
|[Conflicts](Outlook.PostItem.Conflicts.md)|
|[ConversationID](Outlook.PostItem.ConversationID.md)|
|[ConversationIndex](Outlook.PostItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.PostItem.ConversationTopic.md)|
|[CreationTime](Outlook.PostItem.CreationTime.md)|
|[DownloadState](Outlook.PostItem.DownloadState.md)|
|[EntryID](Outlook.PostItem.EntryID.md)|
|[ExpiryTime](Outlook.PostItem.ExpiryTime.md)|
|[FormDescription](Outlook.PostItem.FormDescription.md)|
|[GetInspector](Outlook.PostItem.GetInspector.md)|
|[HTMLBody](Outlook.PostItem.HTMLBody.md)|
|[Importance](Outlook.PostItem.Importance.md)|
|[InternetCodepage](Outlook.PostItem.InternetCodepage.md)|
|[IsConflict](Outlook.PostItem.IsConflict.md)|
|[IsMarkedAsTask](Outlook.PostItem.IsMarkedAsTask.md)|
|[ItemProperties](Outlook.PostItem.ItemProperties.md)|
|[LastModificationTime](Outlook.PostItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.PostItem.MarkForDownload.md)|
|[MessageClass](Outlook.PostItem.MessageClass.md)|
|[Mileage](Outlook.PostItem.Mileage.md)|
|[NoAging](Outlook.PostItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.PostItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.PostItem.OutlookVersion.md)|
|[Parent](Outlook.PostItem.Parent.md)|
|[PropertyAccessor](Outlook.PostItem.PropertyAccessor.md)|
|[ReceivedTime](Outlook.PostItem.ReceivedTime.md)|
|[ReminderOverrideDefault](Outlook.PostItem.ReminderOverrideDefault.md)|
|[ReminderPlaySound](Outlook.PostItem.ReminderPlaySound.md)|
|[ReminderSet](Outlook.PostItem.ReminderSet.md)|
|[ReminderSoundFile](Outlook.PostItem.ReminderSoundFile.md)|
|[ReminderTime](Outlook.PostItem.ReminderTime.md)|
|[RTFBody](Outlook.PostItem.RTFBody.md)|
|[Saved](Outlook.PostItem.Saved.md)|
|[SenderEmailAddress](Outlook.PostItem.SenderEmailAddress.md)|
|[SenderEmailType](Outlook.PostItem.SenderEmailType.md)|
|[SenderName](Outlook.PostItem.SenderName.md)|
|[Sensitivity](Outlook.PostItem.Sensitivity.md)|
|[SentOn](Outlook.PostItem.SentOn.md)|
|[Session](Outlook.PostItem.Session.md)|
|[Size](Outlook.PostItem.Size.md)|
|[Subject](Outlook.PostItem.Subject.md)|
|[TaskCompletedDate](Outlook.PostItem.TaskCompletedDate.md)|
|[TaskDueDate](Outlook.PostItem.TaskDueDate.md)|
|[TaskStartDate](Outlook.PostItem.TaskStartDate.md)|
|[TaskSubject](Outlook.PostItem.TaskSubject.md)|
|[ToDoTaskOrdinal](Outlook.PostItem.ToDoTaskOrdinal.md)|
|[UnRead](Outlook.PostItem.UnRead.md)|
|[UserProperties](Outlook.PostItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]