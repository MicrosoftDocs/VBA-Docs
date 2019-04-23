---
title: JournalItem object (Outlook)
keywords: vbaol11.chm2999
f1_keywords:
- vbaol11.chm2999
ms.prod: outlook
api_name:
- Outlook.JournalItem
ms.assetid: 6e850295-39f9-47b8-e866-9622e9958c69
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem object (Outlook)

Represents a journal entry in a Journal folder. 


## Remarks

A journal entry represents a record of all Outlook-moderated transactions for any given period.

Use the  **[CreateItem](Outlook.Application.CreateItem.md)** method to create a **JournalItem** object that represents a new journal entry.

Use  **[Items](Outlook.Folder.Items.md)** (_index_), where _index_ is the index number of a journal entry or a value used to match the default property of a journal entry, to return a single **JournalItem** object from a Journal folder.


## Example

The following Visual Basic for Applications (VBA) example returns a new journal entry.


```vb
Set myItem = Application.CreateItem(olJournalItem)
```


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.JournalItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.JournalItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.JournalItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.JournalItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.JournalItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.JournalItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.JournalItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.JournalItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.JournalItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.JournalItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.JournalItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.JournalItem.BeforeDelete.md)|
|[BeforeRead](Outlook.JournalItem.BeforeRead.md)|
|[Close](Outlook.JournalItem.Close(even).md)|
|[CustomAction](Outlook.JournalItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.JournalItem.CustomPropertyChange.md)|
|[Forward](Outlook.JournalItem.Forward(even).md)|
|[Open](Outlook.JournalItem.Open.md)|
|[PropertyChange](Outlook.JournalItem.PropertyChange.md)|
|[Read](Outlook.JournalItem.Read.md)|
|[ReadComplete](Outlook.journalitem.readcomplete.md)|
|[Reply](Outlook.JournalItem.Reply(even).md)|
|[ReplyAll](Outlook.JournalItem.ReplyAll(even).md)|
|[Send](Outlook.JournalItem.Send.md)|
|[Unload](Outlook.JournalItem.Unload.md)|
|[Write](Outlook.JournalItem.Write.md)|

## Methods



|Name|
|:-----|
|[Close](Outlook.JournalItem.Close(method).md)|
|[Copy](Outlook.JournalItem.Copy.md)|
|[Delete](Outlook.JournalItem.Delete.md)|
|[Display](Outlook.JournalItem.Display.md)|
|[Forward](Outlook.JournalItem.Forward(method).md)|
|[GetConversation](Outlook.JournalItem.GetConversation.md)|
|[Move](Outlook.JournalItem.Move.md)|
|[PrintOut](Outlook.JournalItem.PrintOut.md)|
|[Reply](Outlook.JournalItem.Reply(method).md)|
|[ReplyAll](Outlook.JournalItem.ReplyAll(method).md)|
|[Save](Outlook.JournalItem.Save.md)|
|[SaveAs](Outlook.JournalItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.JournalItem.ShowCategoriesDialog.md)|
|[StartTimer](Outlook.JournalItem.StartTimer.md)|
|[StopTimer](Outlook.JournalItem.StopTimer.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.JournalItem.Actions.md)|
|[Application](Outlook.JournalItem.Application.md)|
|[Attachments](Outlook.JournalItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.JournalItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.JournalItem.BillingInformation.md)|
|[Body](Outlook.JournalItem.Body.md)|
|[Categories](Outlook.JournalItem.Categories.md)|
|[Class](Outlook.JournalItem.Class.md)|
|[Companies](Outlook.JournalItem.Companies.md)|
|[Conflicts](Outlook.JournalItem.Conflicts.md)|
|[ContactNames](Outlook.JournalItem.ContactNames.md)|
|[ConversationID](Outlook.JournalItem.ConversationID.md)|
|[ConversationIndex](Outlook.JournalItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.JournalItem.ConversationTopic.md)|
|[CreationTime](Outlook.JournalItem.CreationTime.md)|
|[DocPosted](Outlook.JournalItem.DocPosted.md)|
|[DocPrinted](Outlook.JournalItem.DocPrinted.md)|
|[DocRouted](Outlook.JournalItem.DocRouted.md)|
|[DocSaved](Outlook.JournalItem.DocSaved.md)|
|[DownloadState](Outlook.JournalItem.DownloadState.md)|
|[Duration](Outlook.JournalItem.Duration.md)|
|[End](Outlook.JournalItem.End.md)|
|[EntryID](Outlook.JournalItem.EntryID.md)|
|[FormDescription](Outlook.JournalItem.FormDescription.md)|
|[GetInspector](Outlook.JournalItem.GetInspector.md)|
|[Importance](Outlook.JournalItem.Importance.md)|
|[IsConflict](Outlook.JournalItem.IsConflict.md)|
|[ItemProperties](Outlook.JournalItem.ItemProperties.md)|
|[LastModificationTime](Outlook.JournalItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.JournalItem.MarkForDownload.md)|
|[MessageClass](Outlook.JournalItem.MessageClass.md)|
|[Mileage](Outlook.JournalItem.Mileage.md)|
|[NoAging](Outlook.JournalItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.JournalItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.JournalItem.OutlookVersion.md)|
|[Parent](Outlook.JournalItem.Parent.md)|
|[PropertyAccessor](Outlook.JournalItem.PropertyAccessor.md)|
|[Recipients](Outlook.JournalItem.Recipients.md)|
|[Saved](Outlook.JournalItem.Saved.md)|
|[Sensitivity](Outlook.JournalItem.Sensitivity.md)|
|[Session](Outlook.JournalItem.Session.md)|
|[Size](Outlook.JournalItem.Size.md)|
|[Start](Outlook.JournalItem.Start.md)|
|[Subject](Outlook.JournalItem.Subject.md)|
|[Type](Outlook.JournalItem.Type.md)|
|[UnRead](Outlook.JournalItem.UnRead.md)|
|[UserProperties](Outlook.JournalItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]