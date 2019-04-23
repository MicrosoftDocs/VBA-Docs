---
title: TaskRequestDeclineItem object (Outlook)
keywords: vbaol11.chm3009
f1_keywords:
- vbaol11.chm3009
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem
ms.assetid: e842c7c0-7943-9219-329b-30b892ab99b0
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem object (Outlook)

Represents a response to a  **[TaskRequestItem](Outlook.TaskRequestItem.md)** sent by the initiating user.


## Remarks

If the delegated user declines the task, the  **[ResponseState](Outlook.TaskItem.ResponseState.md)** property is set to **olTaskDecline**. The associated **[TaskItem](Outlook.TaskItem.md)** is received by the delegator as a **TaskRequestDeclineItem** object.

Unlike other Microsoft Outlook objects, you cannot create this object.

Use the  **[GetAssociatedTask](Outlook.TaskRequestDeclineItem.GetAssociatedTask.md)** method to return the **TaskItem** object that is associated with this **TaskRequestDeclineItem**. Work directly with the **TaskItem** object.


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.TaskRequestDeclineItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.TaskRequestDeclineItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.TaskRequestDeclineItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.TaskRequestDeclineItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.TaskRequestDeclineItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.TaskRequestDeclineItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.TaskRequestDeclineItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.TaskRequestDeclineItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.TaskRequestDeclineItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.TaskRequestDeclineItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.TaskRequestDeclineItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.TaskRequestDeclineItem.BeforeDelete.md)|
|[BeforeRead](Outlook.TaskRequestDeclineItem.BeforeRead.md)|
|[Close](Outlook.TaskRequestDeclineItem.Close(even).md)|
|[CustomAction](Outlook.TaskRequestDeclineItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.TaskRequestDeclineItem.CustomPropertyChange.md)|
|[Forward](Outlook.TaskRequestDeclineItem.Forward.md)|
|[Open](Outlook.TaskRequestDeclineItem.Open.md)|
|[PropertyChange](Outlook.TaskRequestDeclineItem.PropertyChange.md)|
|[Read](Outlook.TaskRequestDeclineItem.Read.md)|
|[ReadComplete](Outlook.taskrequestdeclineitem.readcomplete.md)|
|[Reply](Outlook.TaskRequestDeclineItem.Reply.md)|
|[ReplyAll](Outlook.TaskRequestDeclineItem.ReplyAll.md)|
|[Send](Outlook.TaskRequestDeclineItem.Send.md)|
|[Unload](Outlook.TaskRequestDeclineItem.Unload.md)|
|[Write](Outlook.TaskRequestDeclineItem.Write.md)|

## Methods



|Name|
|:-----|
|[Close](Outlook.TaskRequestDeclineItem.Close(method).md)|
|[Copy](Outlook.TaskRequestDeclineItem.Copy.md)|
|[Delete](Outlook.TaskRequestDeclineItem.Delete.md)|
|[Display](Outlook.TaskRequestDeclineItem.Display.md)|
|[GetAssociatedTask](Outlook.TaskRequestDeclineItem.GetAssociatedTask.md)|
|[GetConversation](Outlook.TaskRequestDeclineItem.GetConversation.md)|
|[Move](Outlook.TaskRequestDeclineItem.Move.md)|
|[PrintOut](Outlook.TaskRequestDeclineItem.PrintOut.md)|
|[Save](Outlook.TaskRequestDeclineItem.Save.md)|
|[SaveAs](Outlook.TaskRequestDeclineItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.TaskRequestDeclineItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.TaskRequestDeclineItem.Actions.md)|
|[Application](Outlook.TaskRequestDeclineItem.Application.md)|
|[Attachments](Outlook.TaskRequestDeclineItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.TaskRequestDeclineItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.TaskRequestDeclineItem.BillingInformation.md)|
|[Body](Outlook.TaskRequestDeclineItem.Body.md)|
|[Categories](Outlook.TaskRequestDeclineItem.Categories.md)|
|[Class](Outlook.TaskRequestDeclineItem.Class.md)|
|[Companies](Outlook.TaskRequestDeclineItem.Companies.md)|
|[Conflicts](Outlook.TaskRequestDeclineItem.Conflicts.md)|
|[ConversationID](Outlook.TaskRequestDeclineItem.ConversationID.md)|
|[ConversationIndex](Outlook.TaskRequestDeclineItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.TaskRequestDeclineItem.ConversationTopic.md)|
|[CreationTime](Outlook.TaskRequestDeclineItem.CreationTime.md)|
|[DownloadState](Outlook.TaskRequestDeclineItem.DownloadState.md)|
|[EntryID](Outlook.TaskRequestDeclineItem.EntryID.md)|
|[FormDescription](Outlook.TaskRequestDeclineItem.FormDescription.md)|
|[GetInspector](Outlook.TaskRequestDeclineItem.GetInspector.md)|
|[Importance](Outlook.TaskRequestDeclineItem.Importance.md)|
|[IsConflict](Outlook.TaskRequestDeclineItem.IsConflict.md)|
|[ItemProperties](Outlook.TaskRequestDeclineItem.ItemProperties.md)|
|[LastModificationTime](Outlook.TaskRequestDeclineItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.TaskRequestDeclineItem.MarkForDownload.md)|
|[MessageClass](Outlook.TaskRequestDeclineItem.MessageClass.md)|
|[Mileage](Outlook.TaskRequestDeclineItem.Mileage.md)|
|[NoAging](Outlook.TaskRequestDeclineItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.TaskRequestDeclineItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.TaskRequestDeclineItem.OutlookVersion.md)|
|[Parent](Outlook.TaskRequestDeclineItem.Parent.md)|
|[PropertyAccessor](Outlook.TaskRequestDeclineItem.PropertyAccessor.md)|
|[RTFBody](Outlook.TaskRequestDeclineItem.RTFBody.md)|
|[Saved](Outlook.TaskRequestDeclineItem.Saved.md)|
|[Sensitivity](Outlook.TaskRequestDeclineItem.Sensitivity.md)|
|[Session](Outlook.TaskRequestDeclineItem.Session.md)|
|[Size](Outlook.TaskRequestDeclineItem.Size.md)|
|[Subject](Outlook.TaskRequestDeclineItem.Subject.md)|
|[UnRead](Outlook.TaskRequestDeclineItem.UnRead.md)|
|[UserProperties](Outlook.TaskRequestDeclineItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]