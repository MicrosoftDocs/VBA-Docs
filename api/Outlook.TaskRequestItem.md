---
title: TaskRequestItem object (Outlook)
keywords: vbaol11.chm3010
f1_keywords:
- vbaol11.chm3010
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem
ms.assetid: 2908a28a-634c-e786-aa53-f3e32038b727
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem object (Outlook)

Represents a change to the recipient's Tasks list initiated by another party or as a result of a group tasking.


## Remarks

Unlike other Microsoft Outlook objects, you cannot create this object. When the sender applies the  **[Assign](Outlook.TaskItem.Assign.md)** and **[Send](Outlook.TaskItem.Send(method).md)** methods to a **[TaskItem](Outlook.TaskItem.md)** object to assign (delegate) the associated task to another user, the **TaskRequestItem** object is created when the item is received in the recipient's Inbox.

Use the  **[GetAssociatedTask](Outlook.TaskRequestItem.GetAssociatedTask.md)** method to return the **TaskItem** object, and work directly with the **TaskItem** object to respond to the request.


## Example

The following Visual Basic for Applications (VBA) example creates a simple task, assigns it to another user, and sends it. When the task request arrives in the recipient's Inbox, it is received as a **TaskRequestItem**.






```vb
Sub SendTask() 
 
 Dim myItem As Outlook.TaskItem 
 
 Dim myDelegate As Outlook.Recipient 
 
 
 
 Set myItem = Application.CreateItem(olTaskItem) 
 
 myItem.Assign 
 
 Set myDelegate = myItem.Recipients.Add("Jeff Smith") 
 
 myItem.Subject = "Prepare Agenda For Meeting" 
 
 myItem.DueDate = #9/20/97# 
 
 myItem.Send 
 
End Sub
```


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.TaskRequestItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.TaskRequestItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.TaskRequestItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.TaskRequestItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.TaskRequestItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.TaskRequestItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.TaskRequestItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.TaskRequestItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.TaskRequestItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.TaskRequestItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.TaskRequestItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.TaskRequestItem.BeforeDelete.md)|
|[BeforeRead](Outlook.TaskRequestItem.BeforeRead.md)|
|[Close](Outlook.TaskRequestItem.Close(even).md)|
|[CustomAction](Outlook.TaskRequestItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.TaskRequestItem.CustomPropertyChange.md)|
|[Forward](Outlook.TaskRequestItem.Forward.md)|
|[Open](Outlook.TaskRequestItem.Open.md)|
|[PropertyChange](Outlook.TaskRequestItem.PropertyChange.md)|
|[Read](Outlook.TaskRequestItem.Read.md)|
|[ReadComplete](Outlook.taskrequestitem.readcomplete.md)|
|[Reply](Outlook.TaskRequestItem.Reply.md)|
|[ReplyAll](Outlook.TaskRequestItem.ReplyAll.md)|
|[Send](Outlook.TaskRequestItem.Send.md)|
|[Unload](Outlook.TaskRequestItem.Unload.md)|
|[Write](Outlook.TaskRequestItem.Write.md)|

## Methods



|Name|
|:-----|
|[Close](Outlook.TaskRequestItem.Close(method).md)|
|[Copy](Outlook.TaskRequestItem.Copy.md)|
|[Delete](Outlook.TaskRequestItem.Delete.md)|
|[Display](Outlook.TaskRequestItem.Display.md)|
|[GetAssociatedTask](Outlook.TaskRequestItem.GetAssociatedTask.md)|
|[GetConversation](Outlook.TaskRequestItem.GetConversation.md)|
|[Move](Outlook.TaskRequestItem.Move.md)|
|[PrintOut](Outlook.TaskRequestItem.PrintOut.md)|
|[Save](Outlook.TaskRequestItem.Save.md)|
|[SaveAs](Outlook.TaskRequestItem.SaveAs.md)|
|[ShowCategoriesDialog](Outlook.TaskRequestItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.TaskRequestItem.Actions.md)|
|[Application](Outlook.TaskRequestItem.Application.md)|
|[Attachments](Outlook.TaskRequestItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.TaskRequestItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.TaskRequestItem.BillingInformation.md)|
|[Body](Outlook.TaskRequestItem.Body.md)|
|[Categories](Outlook.TaskRequestItem.Categories.md)|
|[Class](Outlook.TaskRequestItem.Class.md)|
|[Companies](Outlook.TaskRequestItem.Companies.md)|
|[Conflicts](Outlook.TaskRequestItem.Conflicts.md)|
|[ConversationID](Outlook.TaskRequestItem.ConversationID.md)|
|[ConversationIndex](Outlook.TaskRequestItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.TaskRequestItem.ConversationTopic.md)|
|[CreationTime](Outlook.TaskRequestItem.CreationTime.md)|
|[DownloadState](Outlook.TaskRequestItem.DownloadState.md)|
|[EntryID](Outlook.TaskRequestItem.EntryID.md)|
|[FormDescription](Outlook.TaskRequestItem.FormDescription.md)|
|[GetInspector](Outlook.TaskRequestItem.GetInspector.md)|
|[Importance](Outlook.TaskRequestItem.Importance.md)|
|[IsConflict](Outlook.TaskRequestItem.IsConflict.md)|
|[ItemProperties](Outlook.TaskRequestItem.ItemProperties.md)|
|[LastModificationTime](Outlook.TaskRequestItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.TaskRequestItem.MarkForDownload.md)|
|[MessageClass](Outlook.TaskRequestItem.MessageClass.md)|
|[Mileage](Outlook.TaskRequestItem.Mileage.md)|
|[NoAging](Outlook.TaskRequestItem.NoAging.md)|
|[OutlookInternalVersion](Outlook.TaskRequestItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.TaskRequestItem.OutlookVersion.md)|
|[Parent](Outlook.TaskRequestItem.Parent.md)|
|[PropertyAccessor](Outlook.TaskRequestItem.PropertyAccessor.md)|
|[RTFBody](Outlook.TaskRequestItem.RTFBody.md)|
|[Saved](Outlook.TaskRequestItem.Saved.md)|
|[Sensitivity](Outlook.TaskRequestItem.Sensitivity.md)|
|[Session](Outlook.TaskRequestItem.Session.md)|
|[Size](Outlook.TaskRequestItem.Size.md)|
|[Subject](Outlook.TaskRequestItem.Subject.md)|
|[UnRead](Outlook.TaskRequestItem.UnRead.md)|
|[UserProperties](Outlook.TaskRequestItem.UserProperties.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]