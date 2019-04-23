---
title: TaskRequestUpdateItem object (Outlook)
keywords: vbaol11.chm3011
f1_keywords:
- vbaol11.chm3011
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem
ms.assetid: 5bc407fe-b3f6-3e46-8b91-e2ed96292cec
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem object (Outlook)

Represents a response to a  **[TaskRequestItem](Outlook.TaskRequestItem.md)** sent by the initiating user.


## Remarks

If the delegated user updates the task by changing properties such as the  **[DueDate](Outlook.TaskItem.DueDate.md)** or the **[Status](Outlook.TaskItem.Status.md)**, and then sends it, the associated **[TaskItem](Outlook.TaskItem.md)** is received by the delegator as a **TaskRequestUpdateItem** object.

Unlike other Microsoft Outlook objects, you cannot create this object.

Use the  **[GetAssociatedTask](Outlook.TaskRequestUpdateItem.GetAssociatedTask.md)** method to return the **TaskItem** object that is associated with this **TaskRequestUpdateItem**. Work directly with the **TaskItem** object


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.TaskRequestUpdateItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.TaskRequestUpdateItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.TaskRequestUpdateItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.TaskRequestUpdateItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.TaskRequestUpdateItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.TaskRequestUpdateItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.TaskRequestUpdateItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.TaskRequestUpdateItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.TaskRequestUpdateItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.TaskRequestUpdateItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.TaskRequestUpdateItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.TaskRequestUpdateItem.BeforeDelete.md)|
|[BeforeRead](Outlook.TaskRequestUpdateItem.BeforeRead.md)|
|[Close](Outlook.TaskRequestUpdateItem.Close(even).md)|
|[CustomAction](Outlook.TaskRequestUpdateItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.TaskRequestUpdateItem.CustomPropertyChange.md)|
|[Forward](Outlook.TaskRequestUpdateItem.Forward.md)|
|[Open](Outlook.TaskRequestUpdateItem.Open.md)|
|[PropertyChange](Outlook.TaskRequestUpdateItem.PropertyChange.md)|
|[Read](Outlook.TaskRequestUpdateItem.Read.md)|
|[ReadComplete](Outlook.taskrequestupdateitem.readcomplete.md)|
|[Reply](Outlook.TaskRequestUpdateItem.Reply.md)|
|[ReplyAll](Outlook.TaskRequestUpdateItem.ReplyAll.md)|
|[Send](Outlook.TaskRequestUpdateItem.Send.md)|
|[Unload](Outlook.TaskRequestUpdateItem.Unload.md)|
|[Write](Outlook.TaskRequestUpdateItem.Write.md)|

## Methods



|Name|
|:-----|
|[Close](Outlook.TaskRequestUpdateItem.Close(method).md)|
|[Copy](Outlook.TaskRequestUpdateItem.Copy.md)|
|[Delete](Outlook.TaskRequestUpdateItem.Delete.md)|
|[Display](Outlook.TaskRequestUpdateItem.Display.md)|
|[GetAssociatedTask](Outlook.TaskRequestUpdateItem.GetAssociatedTask.md)|
|[GetConversation](Outlook.TaskRequestUpdateItem.GetConversation.md)|
|[Move](Outlook.TaskRequestUpdateItem.Move.md)|
|[PrintOut](Outlook.TaskRequestUpdateItem.PrintOut.md)|
|[Save](Outlook.TaskRequestUpdateItem.Save.md)|
|[SaveAs](Outlook.TaskRequestUpdateItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.TaskRequestUpdateItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.TaskRequestUpdateItem.Actions.md)|
|[Application](Outlook.TaskRequestUpdateItem.Application.md)|
|[Attachments](Outlook.TaskRequestUpdateItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.TaskRequestUpdateItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.TaskRequestUpdateItem.BillingInformation.md)|
|[Body](Outlook.TaskRequestUpdateItem.Body.md)|
|[Categories](Outlook.TaskRequestUpdateItem.Categories.md)|
|[Class](Outlook.TaskRequestUpdateItem.Class.md)|
|[Companies](Outlook.TaskRequestUpdateItem.Companies.md)|
|[Conflicts](Outlook.TaskRequestUpdateItem.Conflicts.md)|
|[ConversationID](Outlook.TaskRequestUpdateItem.ConversationID.md)|
|[ConversationIndex](Outlook.TaskRequestUpdateItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.TaskRequestUpdateItem.ConversationTopic.md)|
|[CreationTime](Outlook.TaskRequestUpdateItem.CreationTime.md)|
|[DownloadState](Outlook.TaskRequestUpdateItem.DownloadState.md)|
|[EntryID](Outlook.TaskRequestUpdateItem.EntryID.md)|
|[FormDescription](Outlook.TaskRequestUpdateItem.FormDescription.md)|
|[GetInspector](Outlook.TaskRequestUpdateItem.GetInspector.md)|
|[Importance](Outlook.TaskRequestUpdateItem.Importance.md)|
|[IsConflict](Outlook.TaskRequestUpdateItem.IsConflict.md)|
|[ItemProperties](Outlook.TaskRequestUpdateItem.ItemProperties.md)|
|[LastModificationTime](Outlook.TaskRequestUpdateItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.TaskRequestUpdateItem.MarkForDownload.md)|
|[MessageClass](Outlook.TaskRequestUpdateItem.MessageClass.md)|
|[Mileage](Outlook.TaskRequestUpdateItem.Mileage.md)|
|[NoAging](Outlook.TaskRequestUpdateItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.TaskRequestUpdateItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.TaskRequestUpdateItem.OutlookVersion.md)|
|[Parent](Outlook.TaskRequestUpdateItem.Parent.md)|
|[PropertyAccessor](Outlook.TaskRequestUpdateItem.PropertyAccessor.md)|
|[RTFBody](Outlook.TaskRequestUpdateItem.RTFBody.md)|
|[Saved](Outlook.TaskRequestUpdateItem.Saved.md)|
|[Sensitivity](Outlook.TaskRequestUpdateItem.Sensitivity.md)|
|[Session](Outlook.TaskRequestUpdateItem.Session.md)|
|[Size](Outlook.TaskRequestUpdateItem.Size.md)|
|[Subject](Outlook.TaskRequestUpdateItem.Subject.md)|
|[UnRead](Outlook.TaskRequestUpdateItem.UnRead.md)|
|[UserProperties](Outlook.TaskRequestUpdateItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]