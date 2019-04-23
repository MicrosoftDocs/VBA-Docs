---
title: TaskRequestAcceptItem object (Outlook)
keywords: vbaol11.chm3008
f1_keywords:
- vbaol11.chm3008
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem
ms.assetid: a2905f72-0a67-b07d-7f85-84fe4de17c25
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem object (Outlook)

Represents a response to a  **[TaskRequestItem](Outlook.TaskRequestItem.md)** sent by the initiating user.


## Remarks

If the delegated user accepts the task, the  **[ResponseState](Outlook.TaskItem.ResponseState.md)** property is set to **olTaskAccept**. The associated **[TaskItem](Outlook.TaskItem.md)** is received by the delegator as a **TaskRequestAcceptItem** object.

Unlike other Microsoft Outlook objects, you cannot create this object.

Use the  **[GetAssociatedTask](Outlook.TaskRequestAcceptItem.GetAssociatedTask.md)** method to return the **TaskItem** object that is associated with this **TaskRequestAcceptItem**. Work directly with the **TaskItem** object.


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.TaskRequestAcceptItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.TaskRequestAcceptItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.TaskRequestAcceptItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.TaskRequestAcceptItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.TaskRequestAcceptItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.TaskRequestAcceptItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.TaskRequestAcceptItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.TaskRequestAcceptItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.TaskRequestAcceptItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.TaskRequestAcceptItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.TaskRequestAcceptItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.TaskRequestAcceptItem.BeforeDelete.md)|
|[BeforeRead](Outlook.TaskRequestAcceptItem.BeforeRead.md)|
|[Close](Outlook.TaskRequestAcceptItem.Close(even).md)|
|[CustomAction](Outlook.TaskRequestAcceptItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.TaskRequestAcceptItem.CustomPropertyChange.md)|
|[Forward](Outlook.TaskRequestAcceptItem.Forward.md)|
|[Open](Outlook.TaskRequestAcceptItem.Open.md)|
|[PropertyChange](Outlook.TaskRequestAcceptItem.PropertyChange.md)|
|[Read](Outlook.TaskRequestAcceptItem.Read.md)|
|[ReadComplete](Outlook.taskrequestacceptitem.readcomplete.md)|
|[Reply](Outlook.TaskRequestAcceptItem.Reply.md)|
|[ReplyAll](Outlook.TaskRequestAcceptItem.ReplyAll.md)|
|[Send](Outlook.TaskRequestAcceptItem.Send.md)|
|[Unload](Outlook.TaskRequestAcceptItem.Unload.md)|
|[Write](Outlook.TaskRequestAcceptItem.Write.md)|

## Methods



|Name|
|:-----|
|[Close](Outlook.TaskRequestAcceptItem.Close(method).md)|
|[Copy](Outlook.TaskRequestAcceptItem.Copy.md)|
|[Delete](Outlook.TaskRequestAcceptItem.Delete.md)|
|[Display](Outlook.TaskRequestAcceptItem.Display.md)|
|[GetAssociatedTask](Outlook.TaskRequestAcceptItem.GetAssociatedTask.md)|
|[GetConversation](Outlook.TaskRequestAcceptItem.GetConversation.md)|
|[Move](Outlook.TaskRequestAcceptItem.Move.md)|
|[PrintOut](Outlook.TaskRequestAcceptItem.PrintOut.md)|
|[Save](Outlook.TaskRequestAcceptItem.Save.md)|
|[SaveAs](Outlook.TaskRequestAcceptItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.TaskRequestAcceptItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.TaskRequestAcceptItem.Actions.md)|
|[Application](Outlook.TaskRequestAcceptItem.Application.md)|
|[Attachments](Outlook.TaskRequestAcceptItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.TaskRequestAcceptItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.TaskRequestAcceptItem.BillingInformation.md)|
|[Body](Outlook.TaskRequestAcceptItem.Body.md)|
|[Categories](Outlook.TaskRequestAcceptItem.Categories.md)|
|[Class](Outlook.TaskRequestAcceptItem.Class.md)|
|[Companies](Outlook.TaskRequestAcceptItem.Companies.md)|
|[Conflicts](Outlook.TaskRequestAcceptItem.Conflicts.md)|
|[ConversationID](Outlook.TaskRequestAcceptItem.ConversationID.md)|
|[ConversationIndex](Outlook.TaskRequestAcceptItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.TaskRequestAcceptItem.ConversationTopic.md)|
|[CreationTime](Outlook.TaskRequestAcceptItem.CreationTime.md)|
|[DownloadState](Outlook.TaskRequestAcceptItem.DownloadState.md)|
|[EntryID](Outlook.TaskRequestAcceptItem.EntryID.md)|
|[FormDescription](Outlook.TaskRequestAcceptItem.FormDescription.md)|
|[GetInspector](Outlook.TaskRequestAcceptItem.GetInspector.md)|
|[Importance](Outlook.TaskRequestAcceptItem.Importance.md)|
|[IsConflict](Outlook.TaskRequestAcceptItem.IsConflict.md)|
|[ItemProperties](Outlook.TaskRequestAcceptItem.ItemProperties.md)|
|[LastModificationTime](Outlook.TaskRequestAcceptItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.TaskRequestAcceptItem.MarkForDownload.md)|
|[MessageClass](Outlook.TaskRequestAcceptItem.MessageClass.md)|
|[Mileage](Outlook.TaskRequestAcceptItem.Mileage.md)|
|[NoAging](Outlook.TaskRequestAcceptItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.TaskRequestAcceptItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.TaskRequestAcceptItem.OutlookVersion.md)|
|[Parent](Outlook.TaskRequestAcceptItem.Parent.md)|
|[PropertyAccessor](Outlook.TaskRequestAcceptItem.PropertyAccessor.md)|
|[RTFBody](Outlook.TaskRequestAcceptItem.RTFBody.md)|
|[Saved](Outlook.TaskRequestAcceptItem.Saved.md)|
|[Sensitivity](Outlook.TaskRequestAcceptItem.Sensitivity.md)|
|[Session](Outlook.TaskRequestAcceptItem.Session.md)|
|[Size](Outlook.TaskRequestAcceptItem.Size.md)|
|[Subject](Outlook.TaskRequestAcceptItem.Subject.md)|
|[UnRead](Outlook.TaskRequestAcceptItem.UnRead.md)|
|[UserProperties](Outlook.TaskRequestAcceptItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]