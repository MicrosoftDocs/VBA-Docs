---
title: DistListItem object (Outlook)
keywords: vbaol11.chm2993
f1_keywords:
- vbaol11.chm2993
ms.prod: outlook
api_name:
- Outlook.DistListItem
ms.assetid: 027c3986-abff-d9b1-ecc2-26d60805e952
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem object (Outlook)

Represents a distribution list in a Contacts folder.


## Remarks

 A distribution list can contain multiple recipients and is used to send messages to everyone in the list.

Use the  **[CreateItem](Outlook.Application.CreateItem.md)** method to create a **DistListItem** object that represents a new distribution list.

Use  **[Items](Outlook.Folder.Items.md)** (_index_), where _index_ is the index number of an item in a contacts folder or a value used to match the default property of an item in the folder, to return a single **DistListItem** object from a contacts folder (that is, a folder whose default item type is **olContactItem**).


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates and displays a new distribution list.


```vb
Set myItem = Application.CreateItem(olDistributionListItem) 
 
myItem.Display
```

The following Visual Basic for Applications example sets the current folder as the contacts folder and displays an existing distribution list named Project Team in the folder.




```vb
Set myNamespace = Application.GetNamespace("MAPI") 
 
Set myFolder = myNamespace.GetDefaultFolder(olFolderContacts) 
 
myFolder.Display 
 
Set myItem = myFolder.Items("Project Team") 
 
myItem.Display
```


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.DistListItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.DistListItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.DistListItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.DistListItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.DistListItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.DistListItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.DistListItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.DistListItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.DistListItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.DistListItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.DistListItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.DistListItem.BeforeDelete.md)|
|[BeforeRead](Outlook.DistListItem.BeforeRead.md)|
|[Close](Outlook.DistListItem.Close(even).md)|
|[CustomAction](Outlook.DistListItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.DistListItem.CustomPropertyChange.md)|
|[Forward](Outlook.DistListItem.Forward.md)|
|[Open](Outlook.DistListItem.Open.md)|
|[PropertyChange](Outlook.DistListItem.PropertyChange.md)|
|[Read](Outlook.DistListItem.Read.md)|
|[ReadComplete](Outlook.distlistitem.readcomplete.md)|
|[Reply](Outlook.DistListItem.Reply.md)|
|[ReplyAll](Outlook.DistListItem.ReplyAll.md)|
|[Send](Outlook.DistListItem.Send.md)|
|[Unload](Outlook.DistListItem.Unload.md)|
|[Write](Outlook.DistListItem.Write.md)|

## Methods



|Name|
|:-----|
|[AddMember](Outlook.DistListItem.AddMember.md)|
|[AddMembers](Outlook.DistListItem.AddMembers.md)|
|[ClearTaskFlag](Outlook.DistListItem.ClearTaskFlag.md)|
|[Close](Outlook.DistListItem.Close(method).md)|
|[Copy](Outlook.DistListItem.Copy.md)|
|[Delete](Outlook.DistListItem.Delete.md)|
|[Display](Outlook.DistListItem.Display.md)|
|[GetConversation](Outlook.DistListItem.GetConversation.md)|
|[GetMember](Outlook.DistListItem.GetMember.md)|
|[MarkAsTask](Outlook.DistListItem.MarkAsTask.md)|
|[Move](Outlook.DistListItem.Move.md)|
|[PrintOut](Outlook.DistListItem.PrintOut.md)|
|[RemoveMember](Outlook.DistListItem.RemoveMember.md)|
|[RemoveMembers](Outlook.DistListItem.RemoveMembers.md)|
|[Save](Outlook.DistListItem.Save.md)|
|[SaveAs](Outlook.DistListItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.DistListItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.DistListItem.Actions.md)|
|[Application](Outlook.DistListItem.Application.md)|
|[Attachments](Outlook.DistListItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.DistListItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.DistListItem.BillingInformation.md)|
|[Body](Outlook.DistListItem.Body.md)|
|[Categories](Outlook.DistListItem.Categories.md)|
|[Class](Outlook.DistListItem.Class.md)|
|[Companies](Outlook.DistListItem.Companies.md)|
|[Conflicts](Outlook.DistListItem.Conflicts.md)|
|[ConversationID](Outlook.DistListItem.ConversationID.md)|
|[ConversationIndex](Outlook.DistListItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.DistListItem.ConversationTopic.md)|
|[CreationTime](Outlook.DistListItem.CreationTime.md)|
|[DLName](Outlook.DistListItem.DLName.md)|
|[DownloadState](Outlook.DistListItem.DownloadState.md)|
|[EntryID](Outlook.DistListItem.EntryID.md)|
|[FormDescription](Outlook.DistListItem.FormDescription.md)|
|[GetInspector](Outlook.DistListItem.GetInspector.md)|
|[Importance](Outlook.DistListItem.Importance.md)|
|[IsConflict](Outlook.DistListItem.IsConflict.md)|
|[IsMarkedAsTask](Outlook.DistListItem.IsMarkedAsTask.md)|
|[ItemProperties](Outlook.DistListItem.ItemProperties.md)|
|[LastModificationTime](Outlook.DistListItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.DistListItem.MarkForDownload.md)|
|[MemberCount](Outlook.DistListItem.MemberCount.md)|
|[MessageClass](Outlook.DistListItem.MessageClass.md)|
|[Mileage](Outlook.DistListItem.Mileage.md)|
|[NoAging](Outlook.DistListItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.DistListItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.DistListItem.OutlookVersion.md)|
|[Parent](Outlook.DistListItem.Parent.md)|
|[PropertyAccessor](Outlook.DistListItem.PropertyAccessor.md)|
|[ReminderOverrideDefault](Outlook.DistListItem.ReminderOverrideDefault.md)|
|[ReminderPlaySound](Outlook.DistListItem.ReminderPlaySound.md)|
|[ReminderSet](Outlook.DistListItem.ReminderSet.md)|
|[ReminderSoundFile](Outlook.DistListItem.ReminderSoundFile.md)|
|[ReminderTime](Outlook.DistListItem.ReminderTime.md)|
|[RTFBody](Outlook.DistListItem.RTFBody.md)|
|[Saved](Outlook.DistListItem.Saved.md)|
|[Sensitivity](Outlook.DistListItem.Sensitivity.md)|
|[Session](Outlook.DistListItem.Session.md)|
|[Size](Outlook.DistListItem.Size.md)|
|[Subject](Outlook.DistListItem.Subject.md)|
|[TaskCompletedDate](Outlook.DistListItem.TaskCompletedDate.md)|
|[TaskDueDate](Outlook.DistListItem.TaskDueDate.md)|
|[TaskStartDate](Outlook.DistListItem.TaskStartDate.md)|
|[TaskSubject](Outlook.DistListItem.TaskSubject.md)|
|[ToDoTaskOrdinal](Outlook.DistListItem.ToDoTaskOrdinal.md)|
|[UnRead](Outlook.DistListItem.UnRead.md)|
|[UserProperties](Outlook.DistListItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]