---
title: DocumentItem Object (Outlook)
keywords: vbaol11.chm2994
f1_keywords:
- vbaol11.chm2994
ms.prod: outlook
api_name:
- Outlook.DocumentItem
ms.assetid: 7b0a6af0-6632-3ff6-841f-5b081d0d68d8
ms.date: 06/08/2017
---


# DocumentItem Object (Outlook)

Represents any document other than a Microsoft Outlook item as an item in an Outlook folder. 


## Remarks

A  **DocumentItem** object is any document other than an Outlook item as an item in an Outlook folder. In common usage, this will be an Office document but may be any type of document or executable file.

Unlike other Outlook objects, you cannot create this object.


 **Note**  When you try to programmatically add a user-defined property to a  **DocumentItem** object, you receive the following error message: "Property is read-only." This is because the Outlook object model does not support this functionality.


## Events



|**Name**|
|:-----|
|[AfterWrite](../../../api/Outlook.DocumentItem.AfterWrite.md)|
|[AttachmentAdd](../../../api/Outlook.DocumentItem.AttachmentAdd.md)|
|[AttachmentRead](../../../api/Outlook.DocumentItem.AttachmentRead.md)|
|[AttachmentRemove](../../../api/Outlook.DocumentItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](../../../api/Outlook.DocumentItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](../../../api/Outlook.DocumentItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](../../../api/Outlook.DocumentItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](../../../api/Outlook.DocumentItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](../../../api/Outlook.DocumentItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](../../../api/Outlook.DocumentItem.BeforeAutoSave.md)|
|[BeforeCheckNames](../../../api/Outlook.DocumentItem.BeforeCheckNames.md)|
|[BeforeDelete](../../../api/Outlook.DocumentItem.BeforeDelete.md)|
|[BeforeRead](../../../api/Outlook.DocumentItem.BeforeRead.md)|
|[Close](../../../api/Outlook.DocumentItem.Close(even).md)|
|[CustomAction](../../../api/Outlook.DocumentItem.CustomAction.md)|
|[CustomPropertyChange](../../../api/Outlook.DocumentItem.CustomPropertyChange.md)|
|[Forward](../../../api/Outlook.DocumentItem.Forward.md)|
|[Open](../../../api/Outlook.DocumentItem.Open.md)|
|[PropertyChange](../../../api/Outlook.DocumentItem.PropertyChange.md)|
|[Read](../../../api/Outlook.DocumentItem.Read.md)|
|[ReadComplete](../../../api/Outlook.documentitem.readcomplete.md)|
|[Reply](../../../api/Outlook.DocumentItem.Reply.md)|
|[ReplyAll](../../../api/Outlook.DocumentItem.ReplyAll.md)|
|[Send](../../../api/Outlook.DocumentItem.Send.md)|
|[Unload](../../../api/Outlook.DocumentItem.Unload.md)|
|[Write](../../../api/Outlook.DocumentItem.Write.md)|

## Methods



|**Name**|
|:-----|
|[Close](../../../api/Outlook.DocumentItem.Close(method).md)|
|[Copy](../../../api/Outlook.DocumentItem.Copy.md)|
|[Delete](../../../api/Outlook.DocumentItem.Delete.md)|
|[Display](../../../api/Outlook.DocumentItem.Display.md)|
|[Move](../../../api/Outlook.DocumentItem.Move.md)|
|[PrintOut](../../../api/Outlook.DocumentItem.PrintOut.md)|
|[Save](../../../api/Outlook.DocumentItem.Save.md)|
|[SaveAs](../../../api/Outlook.DocumentItem.SaveAs.md)|
|[ShowCategoriesDialog](../../../api/Outlook.DocumentItem.ShowCategoriesDialog.md)|

## Properties



|**Name**|
|:-----|
|[Actions](../../../api/Outlook.DocumentItem.Actions.md)|
|[Application](../../../api/Outlook.DocumentItem.Application.md)|
|[Attachments](../../../api/Outlook.DocumentItem.Attachments.md)|
|[AutoResolvedWinner](../../../api/Outlook.DocumentItem.AutoResolvedWinner.md)|
|[BillingInformation](../../../api/Outlook.DocumentItem.BillingInformation.md)|
|[Body](../../../api/Outlook.DocumentItem.Body.md)|
|[Categories](../../../api/Outlook.DocumentItem.Categories.md)|
|[Class](../../../api/Outlook.DocumentItem.Class.md)|
|[Companies](../../../api/Outlook.DocumentItem.Companies.md)|
|[Conflicts](../../../api/Outlook.DocumentItem.Conflicts.md)|
|[ConversationIndex](../../../api/Outlook.DocumentItem.ConversationIndex.md)|
|[ConversationTopic](../../../api/Outlook.DocumentItem.ConversationTopic.md)|
|[CreationTime](../../../api/Outlook.DocumentItem.CreationTime.md)|
|[DownloadState](../../../api/Outlook.DocumentItem.DownloadState.md)|
|[EntryID](../../../api/Outlook.DocumentItem.EntryID.md)|
|[FormDescription](../../../api/Outlook.DocumentItem.FormDescription.md)|
|[GetInspector](../../../api/Outlook.DocumentItem.GetInspector.md)|
|[Importance](../../../api/Outlook.DocumentItem.Importance.md)|
|[IsConflict](../../../api/Outlook.DocumentItem.IsConflict.md)|
|[ItemProperties](../../../api/Outlook.DocumentItem.ItemProperties.md)|
|[LastModificationTime](../../../api/Outlook.DocumentItem.LastModificationTime.md)|
|[MarkForDownload](../../../api/Outlook.DocumentItem.MarkForDownload.md)|
|[MessageClass](../../../api/Outlook.DocumentItem.MessageClass.md)|
|[Mileage](../../../api/Outlook.DocumentItem.Mileage.md)|
|[NoAging](../../../api/Outlook.DocumentItem.NoAging.md)|
|[OutlookInternalVersion](../../../api/Outlook.DocumentItem.OutlookInternalVersion.md)|
|[OutlookVersion](../../../api/Outlook.DocumentItem.OutlookVersion.md)|
|[Parent](../../../api/Outlook.DocumentItem.Parent.md)|
|[PropertyAccessor](../../../api/Outlook.DocumentItem.PropertyAccessor.md)|
|[Saved](../../../api/Outlook.DocumentItem.Saved.md)|
|[Sensitivity](../../../api/Outlook.DocumentItem.Sensitivity.md)|
|[Session](../../../api/Outlook.DocumentItem.Session.md)|
|[Size](../../../api/Outlook.DocumentItem.Size.md)|
|[Subject](../../../api/Outlook.DocumentItem.Subject.md)|
|[UnRead](../../../api/Outlook.DocumentItem.UnRead.md)|
|[UserProperties](../../../api/Outlook.DocumentItem.UserProperties.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
