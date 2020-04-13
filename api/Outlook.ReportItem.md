---
title: ReportItem object (Outlook)
keywords: vbaol11.chm3007
f1_keywords:
- vbaol11.chm3007
ms.prod: outlook
api_name:
- Outlook.ReportItem
ms.assetid: 16ebe336-72e0-42f6-99d3-edecc3ea284d
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportItem object (Outlook)

Represents a mail-delivery report in an Inbox folder. 


## Remarks

The **ReportItem** object is similar to a **[MailItem](Outlook.MailItem.md)** object, and it contains a report (usually the non-delivery report) or error message from the mail transport system.

Unlike other Microsoft Outlook objects, you cannot create this object. Report items are created automatically when any report or error in general is received from the mail transport system.


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.ReportItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.ReportItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.ReportItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.ReportItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.ReportItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.ReportItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.ReportItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.ReportItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.ReportItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.ReportItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.ReportItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.ReportItem.BeforeDelete.md)|
|[BeforeRead](Outlook.ReportItem.BeforeRead.md)|
|[Close](Outlook.ReportItem.Close(even).md)|
|[CustomAction](Outlook.ReportItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.ReportItem.CustomPropertyChange.md)|
|[Forward](Outlook.ReportItem.Forward.md)|
|[Open](Outlook.ReportItem.Open.md)|
|[PropertyChange](Outlook.ReportItem.PropertyChange.md)|
|[Read](Outlook.ReportItem.Read.md)|
|[ReadComplete](Outlook.reportitem.readcomplete.md)|
|[Reply](Outlook.ReportItem.Reply.md)|
|[ReplyAll](Outlook.ReportItem.ReplyAll.md)|
|[Send](Outlook.ReportItem.Send.md)|
|[Unload](Outlook.ReportItem.Unload.md)|
|[Write](Outlook.ReportItem.Write.md)|

## Methods



|Name|
|:-----|
|[Close](Outlook.ReportItem.Close(method).md)|
|[Copy](Outlook.ReportItem.Copy.md)|
|[Delete](Outlook.ReportItem.Delete.md)|
|[Display](Outlook.ReportItem.Display.md)|
|[GetConversation](Outlook.ReportItem.GetConversation.md)|
|[Move](Outlook.ReportItem.Move.md)|
|[PrintOut](Outlook.ReportItem.PrintOut.md)|
|[Save](Outlook.ReportItem.Save.md)|
|[SaveAs](Outlook.ReportItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.ReportItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.ReportItem.Actions.md)|
|[Application](Outlook.ReportItem.Application.md)|
|[Attachments](Outlook.ReportItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.ReportItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.ReportItem.BillingInformation.md)|
|[Body](Outlook.ReportItem.Body.md)|
|[Categories](Outlook.ReportItem.Categories.md)|
|[Class](Outlook.ReportItem.Class.md)|
|[Companies](Outlook.ReportItem.Companies.md)|
|[Conflicts](Outlook.ReportItem.Conflicts.md)|
|[ConversationID](Outlook.ReportItem.ConversationID.md)|
|[ConversationIndex](Outlook.ReportItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.ReportItem.ConversationTopic.md)|
|[CreationTime](Outlook.ReportItem.CreationTime.md)|
|[DownloadState](Outlook.ReportItem.DownloadState.md)|
|[EntryID](Outlook.ReportItem.EntryID.md)|
|[FormDescription](Outlook.ReportItem.FormDescription.md)|
|[GetInspector](Outlook.ReportItem.GetInspector.md)|
|[Importance](Outlook.ReportItem.Importance.md)|
|[IsConflict](Outlook.ReportItem.IsConflict.md)|
|[ItemProperties](Outlook.ReportItem.ItemProperties.md)|
|[LastModificationTime](Outlook.ReportItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.ReportItem.MarkForDownload.md)|
|[MessageClass](Outlook.ReportItem.MessageClass.md)|
|[Mileage](Outlook.ReportItem.Mileage.md)|
|[NoAging](Outlook.ReportItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.ReportItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.ReportItem.OutlookVersion.md)|
|[Parent](Outlook.ReportItem.Parent.md)|
|[PropertyAccessor](Outlook.ReportItem.PropertyAccessor.md)|
|[RetentionExpirationDate](Outlook.ReportItem.RetentionExpirationDate.md)|
|[RetentionPolicyName](Outlook.ReportItem.RetentionPolicyName.md)|
|[Saved](Outlook.ReportItem.Saved.md)|
|[Sensitivity](Outlook.ReportItem.Sensitivity.md)|
|[Session](Outlook.ReportItem.Session.md)|
|[Size](Outlook.ReportItem.Size.md)|
|[Subject](Outlook.ReportItem.Subject.md)|
|[UnRead](Outlook.ReportItem.UnRead.md)|
|[UserProperties](Outlook.ReportItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[ReportItem Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]