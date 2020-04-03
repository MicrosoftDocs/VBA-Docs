---
title: DocumentItem object (Outlook)
keywords: vbaol11.chm2994
f1_keywords:
- vbaol11.chm2994
ms.prod: outlook
api_name:
- Outlook.DocumentItem
ms.assetid: 7b0a6af0-6632-3ff6-841f-5b081d0d68d8
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem object (Outlook)

Represents any document other than a Microsoft Outlook item as an item in an Outlook folder. 


## Remarks

A  **DocumentItem** object is any document other than an Outlook item as an item in an Outlook folder. In common usage, this will be an Office document but may be any type of document or executable file.

Unlike other Outlook objects, you cannot create this object.


> [!NOTE] 
> When you try to programmatically add a user-defined property to a  **DocumentItem** object, you receive the following error message: "Property is read-only." This is because the Outlook object model does not support this functionality.


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.DocumentItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.DocumentItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.DocumentItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.DocumentItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.DocumentItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.DocumentItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.DocumentItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.DocumentItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.DocumentItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.DocumentItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.DocumentItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.DocumentItem.BeforeDelete.md)|
|[BeforeRead](Outlook.DocumentItem.BeforeRead.md)|
|[Close](Outlook.DocumentItem.Close(even).md)|
|[CustomAction](Outlook.DocumentItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.DocumentItem.CustomPropertyChange.md)|
|[Forward](Outlook.DocumentItem.Forward.md)|
|[Open](Outlook.DocumentItem.Open.md)|
|[PropertyChange](Outlook.DocumentItem.PropertyChange.md)|
|[Read](Outlook.DocumentItem.Read.md)|
|[ReadComplete](Outlook.documentitem.readcomplete.md)|
|[Reply](Outlook.DocumentItem.Reply.md)|
|[ReplyAll](Outlook.DocumentItem.ReplyAll.md)|
|[Send](Outlook.DocumentItem.Send.md)|
|[Unload](Outlook.DocumentItem.Unload.md)|
|[Write](Outlook.DocumentItem.Write.md)|

## Methods



|Name|
|:-----|
|[Close](Outlook.DocumentItem.Close(method).md)|
|[Copy](Outlook.DocumentItem.Copy.md)|
|[Delete](Outlook.DocumentItem.Delete.md)|
|[Display](Outlook.DocumentItem.Display.md)|
|[Move](Outlook.DocumentItem.Move.md)|
|[PrintOut](Outlook.DocumentItem.PrintOut.md)|
|[Save](Outlook.DocumentItem.Save.md)|
|[SaveAs](Outlook.DocumentItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.DocumentItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.DocumentItem.Actions.md)|
|[Application](Outlook.DocumentItem.Application.md)|
|[Attachments](Outlook.DocumentItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.DocumentItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.DocumentItem.BillingInformation.md)|
|[Body](Outlook.DocumentItem.Body.md)|
|[Categories](Outlook.DocumentItem.Categories.md)|
|[Class](Outlook.DocumentItem.Class.md)|
|[Companies](Outlook.DocumentItem.Companies.md)|
|[Conflicts](Outlook.DocumentItem.Conflicts.md)|
|[ConversationIndex](Outlook.DocumentItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.DocumentItem.ConversationTopic.md)|
|[CreationTime](Outlook.DocumentItem.CreationTime.md)|
|[DownloadState](Outlook.DocumentItem.DownloadState.md)|
|[EntryID](Outlook.DocumentItem.EntryID.md)|
|[FormDescription](Outlook.DocumentItem.FormDescription.md)|
|[GetInspector](Outlook.DocumentItem.GetInspector.md)|
|[Importance](Outlook.DocumentItem.Importance.md)|
|[IsConflict](Outlook.DocumentItem.IsConflict.md)|
|[ItemProperties](Outlook.DocumentItem.ItemProperties.md)|
|[LastModificationTime](Outlook.DocumentItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.DocumentItem.MarkForDownload.md)|
|[MessageClass](Outlook.DocumentItem.MessageClass.md)|
|[Mileage](Outlook.DocumentItem.Mileage.md)|
|[NoAging](Outlook.DocumentItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.DocumentItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.DocumentItem.OutlookVersion.md)|
|[Parent](Outlook.DocumentItem.Parent.md)|
|[PropertyAccessor](Outlook.DocumentItem.PropertyAccessor.md)|
|[Saved](Outlook.DocumentItem.Saved.md)|
|[Sensitivity](Outlook.DocumentItem.Sensitivity.md)|
|[Session](Outlook.DocumentItem.Session.md)|
|[Size](Outlook.DocumentItem.Size.md)|
|[Subject](Outlook.DocumentItem.Subject.md)|
|[UnRead](Outlook.DocumentItem.UnRead.md)|
|[UserProperties](Outlook.DocumentItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]