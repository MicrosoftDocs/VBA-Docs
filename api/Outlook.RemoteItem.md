---
title: RemoteItem object (Outlook)
keywords: vbaol11.chm3006
f1_keywords:
- vbaol11.chm3006
ms.prod: outlook
api_name:
- Outlook.RemoteItem
ms.assetid: 6302aaff-cdcf-4d86-60f1-4bed15540d9f
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem object (Outlook)

Represents a remote item in an Inbox folder.


## Remarks

The  **RemoteItem** object is similar to the **[MailItem](Outlook.MailItem.md)** object, but it contains only the **Subject**,  **Received Date** and **Time**,  **Sender**,  **Size**, and the first 256 characters of the body of the message. It is used to give someone connecting in remote mode enough information to decide whether or not to download the corresponding mail message. However, the headers in items contained in an Offline Folders file (.ost) cannot be accessed using the  **RemoteItem** object.

Unlike other Microsoft Outlook objects, you cannot create this object. Remote items are created by Outlook automatically when you use a Remote Access System (RAS) connection. Each  **RemoteItem** object created on the local system corresponds to a preexisting **MailItem** object on the remote system.

The  **RemoteItem** object inherits a number of properties, methods, and events that, because of the nature of the object, have no function. The **Object Browser** shows these properties, methods, and events as belonging to the **RemoteItem** object, but trying to use them will produce no effect.

The methods that do not work for the  **RemoteItem** object include **Close**, **Copy**, **Display**, **Move**, and **Save**.

The properties that do not work for the  **RemoteItem** object include **BillingInformation**, **Body**, **Categories**, **Companies**, and **Mileage**.

The events that do not work for the  **RemoteItem** object include **Open**, **Close**, **Forward**, **Reply**, **ReplyAll**, and **Send**.


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.RemoteItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.RemoteItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.RemoteItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.RemoteItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.RemoteItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.RemoteItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.RemoteItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.RemoteItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.RemoteItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.RemoteItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.RemoteItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.RemoteItem.BeforeDelete.md)|
|[BeforeRead](Outlook.RemoteItem.BeforeRead.md)|
|[Close](Outlook.RemoteItem.Close(even).md)|
|[CustomAction](Outlook.RemoteItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.RemoteItem.CustomPropertyChange.md)|
|[Forward](Outlook.RemoteItem.Forward.md)|
|[Open](Outlook.RemoteItem.Open.md)|
|[PropertyChange](Outlook.RemoteItem.PropertyChange.md)|
|[Read](Outlook.RemoteItem.Read.md)|
|[ReadComplete](Outlook.remoteitem.readcomplete.md)|
|[Reply](Outlook.RemoteItem.Reply.md)|
|[ReplyAll](Outlook.RemoteItem.ReplyAll.md)|
|[Send](Outlook.RemoteItem.Send.md)|
|[Unload](Outlook.RemoteItem.Unload.md)|
|[Write](Outlook.RemoteItem.Write.md)|

## Methods



|Name|
|:-----|
|[Close](Outlook.RemoteItem.Close(method).md)|
|[Copy](Outlook.RemoteItem.Copy.md)|
|[Delete](Outlook.RemoteItem.Delete.md)|
|[Display](Outlook.RemoteItem.Display.md)|
|[GetConversation](Outlook.RemoteItem.GetConversation.md)|
|[Move](Outlook.RemoteItem.Move.md)|
|[PrintOut](Outlook.RemoteItem.PrintOut.md)|
|[Save](Outlook.RemoteItem.Save.md)|
|[SaveAs](Outlook.RemoteItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.RemoteItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.RemoteItem.Actions.md)|
|[Application](Outlook.RemoteItem.Application.md)|
|[Attachments](Outlook.RemoteItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.RemoteItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.RemoteItem.BillingInformation.md)|
|[Body](Outlook.RemoteItem.Body.md)|
|[Categories](Outlook.RemoteItem.Categories.md)|
|[Class](Outlook.RemoteItem.Class.md)|
|[Companies](Outlook.RemoteItem.Companies.md)|
|[Conflicts](Outlook.RemoteItem.Conflicts.md)|
|[ConversationID](Outlook.RemoteItem.ConversationID.md)|
|[ConversationIndex](Outlook.RemoteItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.RemoteItem.ConversationTopic.md)|
|[CreationTime](Outlook.RemoteItem.CreationTime.md)|
|[DownloadState](Outlook.RemoteItem.DownloadState.md)|
|[EntryID](Outlook.RemoteItem.EntryID.md)|
|[FormDescription](Outlook.RemoteItem.FormDescription.md)|
|[GetInspector](Outlook.RemoteItem.GetInspector.md)|
|[HasAttachment](Outlook.RemoteItem.HasAttachment.md)|
|[Importance](Outlook.RemoteItem.Importance.md)|
|[IsConflict](Outlook.RemoteItem.IsConflict.md)|
|[ItemProperties](Outlook.RemoteItem.ItemProperties.md)|
|[LastModificationTime](Outlook.RemoteItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.RemoteItem.MarkForDownload.md)|
|[MessageClass](Outlook.RemoteItem.MessageClass.md)|
|[Mileage](Outlook.RemoteItem.Mileage.md)|
|[NoAging](Outlook.RemoteItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.RemoteItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.RemoteItem.OutlookVersion.md)|
|[Parent](Outlook.RemoteItem.Parent.md)|
|[PropertyAccessor](Outlook.RemoteItem.PropertyAccessor.md)|
|[RemoteMessageClass](Outlook.RemoteItem.RemoteMessageClass.md)|
|[Saved](Outlook.RemoteItem.Saved.md)|
|[Sensitivity](Outlook.RemoteItem.Sensitivity.md)|
|[Session](Outlook.RemoteItem.Session.md)|
|[Size](Outlook.RemoteItem.Size.md)|
|[Subject](Outlook.RemoteItem.Subject.md)|
|[TransferSize](Outlook.RemoteItem.TransferSize.md)|
|[TransferTime](Outlook.RemoteItem.TransferTime.md)|
|[UnRead](Outlook.RemoteItem.UnRead.md)|
|[UserProperties](Outlook.RemoteItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]