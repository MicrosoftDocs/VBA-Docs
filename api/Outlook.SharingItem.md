---
title: SharingItem object (Outlook)
keywords: vbaol11.chm3016
f1_keywords:
- vbaol11.chm3016
ms.prod: outlook
api_name:
- Outlook.SharingItem
ms.assetid: 63dd3451-44f3-7cc4-c6e2-7dad5835a7d2
ms.date: 04/17/2019
localization_priority: Normal
---


# SharingItem object (Outlook)

Represents a sharing message in an Inbox folder.


## Remarks

Use the **[CreateSharingItem](Outlook.NameSpace.CreateSharingItem.md)** method of the **NameSpace** object to create a **SharingItem** object that represents a new sharing request or sharing invitation.

Use **[Item](Outlook.Folders.Item.md)** (_index_), where _index_ is the index number of a sharing message or a value used to match the default property of a message, to return a single **SharingItem** object from an Inbox folder.


## Example

The following Visual Basic for Applications (VBA) example creates and displays a new sharing invitation for the Tasks folder.

```vb
Public Sub CreateTasksSharingItem() 
 On Error GoTo ErrRoutine 
 
 Dim mapiNamespace As Outlook.NameSpace 
 Set mapiNamespace = Outlook.Application.GetNamespace("MAPI") 
 
 Dim tasksFolder As Outlook.Folder 
 Set tasksFolder = mapiNamespace.GetDefaultFolder(Outlook.olFolderTasks) 
 
 Dim invitation As Outlook.SharingItem  
 Set invitation = appNamespace.CreateSharingItem(tasksFolder) 
 
 invitation.Display 
  
EndRoutine:  
 Exit Sub 
  
ErrRoutine: 
 MsgBox Err.Description, vbOKOnly, Err.Number & " - " & Err.Source  
 Resume EndRoutine 
 
End Sub 
```


## Events

- [AfterWrite](Outlook.SharingItem.AfterWrite.md)
- [AttachmentAdd](Outlook.SharingItem.AttachmentAdd.md)
- [AttachmentRead](Outlook.SharingItem.AttachmentRead.md)
- [AttachmentRemove](Outlook.SharingItem.AttachmentRemove.md)
- [BeforeAttachmentAdd](Outlook.SharingItem.BeforeAttachmentAdd.md)
- [BeforeAttachmentPreview](Outlook.SharingItem.BeforeAttachmentPreview.md)
- [BeforeAttachmentRead](Outlook.SharingItem.BeforeAttachmentRead.md)
- [BeforeAttachmentSave](Outlook.SharingItem.BeforeAttachmentSave.md)
- [BeforeAttachmentWriteToTempFile](Outlook.SharingItem.BeforeAttachmentWriteToTempFile.md)
- [BeforeAutoSave](Outlook.SharingItem.BeforeAutoSave.md)
- [BeforeCheckNames](Outlook.SharingItem.BeforeCheckNames.md)
- [BeforeDelete](Outlook.SharingItem.BeforeDelete.md)
- [BeforeRead](Outlook.SharingItem.BeforeRead.md)
- [Close](Outlook.SharingItem.Close(even).md)
- [CustomAction](Outlook.SharingItem.CustomAction.md)
- [CustomPropertyChange](Outlook.SharingItem.CustomPropertyChange.md)
- [Forward](Outlook.SharingItem.Forward(even).md)
- [Open](Outlook.SharingItem.Open.md)
- [PropertyChange](Outlook.SharingItem.PropertyChange.md)
- [Read](Outlook.SharingItem.Read.md)
- [ReadComplete](Outlook.sharingitem.readcomplete.md)
- [Reply](Outlook.SharingItem.Reply(even).md)
- [ReplyAll](Outlook.SharingItem.ReplyAll(even).md)
- [Send](Outlook.SharingItem.Send(even).md)
- [Unload](Outlook.SharingItem.Unload.md)
- [Write](Outlook.SharingItem.Write.md)

## Methods

- [AddBusinessCard](Outlook.SharingItem.AddBusinessCard.md)
- [Allow](Outlook.SharingItem.Allow.md)
- [ClearConversationIndex](Outlook.SharingItem.ClearConversationIndex.md)
- [ClearTaskFlag](Outlook.SharingItem.ClearTaskFlag.md)
- [Close](Outlook.SharingItem.Close(method).md)
- [Copy](Outlook.SharingItem.Copy.md)
- [Delete](Outlook.SharingItem.Delete.md)
- [Deny](Outlook.SharingItem.Deny.md)
- [Display](Outlook.SharingItem.Display.md)
- [Forward](Outlook.SharingItem.Forward(method).md)
- [GetConversation](Outlook.SharingItem.GetConversation.md)
- [MarkAsTask](Outlook.SharingItem.MarkAsTask.md)
- [Move](Outlook.SharingItem.Move.md)
- [OpenSharedFolder](Outlook.SharingItem.OpenSharedFolder.md)
- [PrintOut](Outlook.SharingItem.PrintOut.md)
- [Reply](Outlook.SharingItem.Reply(method).md)
- [ReplyAll](Outlook.SharingItem.ReplyAll(method).md)
- [Save](Outlook.SharingItem.Save.md)
- [SaveAs](Outlook.SharingItem.SaveAs.md)
- [Send](Outlook.SharingItem.Send(method).md)
- [ShowCategoriesDialog](Outlook.SharingItem.ShowCategoriesDialog.md)

## Properties

- [Actions](Outlook.SharingItem.Actions.md)
- [AllowWriteAccess](Outlook.SharingItem.AllowWriteAccess.md)
- [AlternateRecipientAllowed](Outlook.SharingItem.AlternateRecipientAllowed.md)
- [Application](Outlook.SharingItem.Application.md)
- [Attachments](Outlook.SharingItem.Attachments.md)
- [AutoForwarded](Outlook.SharingItem.AutoForwarded.md)
- [BCC](Outlook.SharingItem.BCC.md)
- [BillingInformation](Outlook.SharingItem.BillingInformation.md)
- [Body](Outlook.SharingItem.Body.md)
- [BodyFormat](Outlook.SharingItem.BodyFormat.md)
- [Categories](Outlook.SharingItem.Categories.md)
- [CC](Outlook.SharingItem.CC.md)
- [Class](Outlook.SharingItem.Class.md)
- [Companies](Outlook.SharingItem.Companies.md)
- [Conflicts](Outlook.SharingItem.Conflicts.md)
- [ConversationID](Outlook.SharingItem.ConversationID.md)
- [ConversationIndex](Outlook.SharingItem.ConversationIndex.md)
- [ConversationTopic](Outlook.SharingItem.ConversationTopic.md)
- [CreationTime](Outlook.SharingItem.CreationTime.md)
- [DeferredDeliveryTime](Outlook.SharingItem.DeferredDeliveryTime.md)
- [DeleteAfterSubmit](Outlook.SharingItem.DeleteAfterSubmit.md)
- [DownloadState](Outlook.SharingItem.DownloadState.md)
- [EntryID](Outlook.SharingItem.EntryID.md)
- [ExpiryTime](Outlook.SharingItem.ExpiryTime.md)
- [FlagRequest](Outlook.SharingItem.FlagRequest.md)
- [FormDescription](Outlook.SharingItem.FormDescription.md)
- [GetInspector](Outlook.SharingItem.GetInspector.md)
- [HTMLBody](Outlook.SharingItem.HTMLBody.md)
- [Importance](Outlook.SharingItem.Importance.md)
- [InternetCodepage](Outlook.SharingItem.InternetCodepage.md)
- [IsConflict](Outlook.SharingItem.IsConflict.md)
- [IsMarkedAsTask](Outlook.SharingItem.IsMarkedAsTask.md)
- [ItemProperties](Outlook.SharingItem.ItemProperties.md)
- [LastModificationTime](Outlook.SharingItem.LastModificationTime.md)
- [MarkForDownload](Outlook.SharingItem.MarkForDownload.md)
- [MessageClass](Outlook.SharingItem.MessageClass.md)
- [Mileage](Outlook.SharingItem.Mileage.md)
- [NoAging](Outlook.SharingItem.NoAging.md)
- [OriginatorDeliveryReportRequested](Outlook.SharingItem.OriginatorDeliveryReportRequested.md)
- [OutlookInternalVersion](Outlook.SharingItem.OutlookInternalVersion.md)
- [OutlookVersion](Outlook.SharingItem.OutlookVersion.md)
- [Parent](Outlook.SharingItem.Parent.md)
- [Permission](Outlook.SharingItem.Permission.md)
- [PermissionService](Outlook.SharingItem.PermissionService.md)
- [PermissionTemplateGuid](Outlook.SharingItem.PermissionTemplateGuid.md)
- [PropertyAccessor](Outlook.SharingItem.PropertyAccessor.md)
- [ReadReceiptRequested](Outlook.SharingItem.ReadReceiptRequested.md)
- [ReceivedByEntryID](Outlook.SharingItem.ReceivedByEntryID.md)
- [ReceivedByName](Outlook.SharingItem.ReceivedByName.md)
- [ReceivedOnBehalfOfEntryID](Outlook.SharingItem.ReceivedOnBehalfOfEntryID.md)
- [ReceivedOnBehalfOfName](Outlook.SharingItem.ReceivedOnBehalfOfName.md)
- [ReceivedTime](Outlook.SharingItem.ReceivedTime.md)
- [RecipientReassignmentProhibited](Outlook.SharingItem.RecipientReassignmentProhibited.md)
- [Recipients](Outlook.SharingItem.Recipients.md)
- [ReminderOverrideDefault](Outlook.SharingItem.ReminderOverrideDefault.md)
- [ReminderPlaySound](Outlook.SharingItem.ReminderPlaySound.md)
- [ReminderSet](Outlook.SharingItem.ReminderSet.md)
- [ReminderSoundFile](Outlook.SharingItem.ReminderSoundFile.md)
- [ReminderTime](Outlook.SharingItem.ReminderTime.md)
- [RemoteID](Outlook.SharingItem.RemoteID.md)
- [RemoteName](Outlook.SharingItem.RemoteName.md)
- [RemotePath](Outlook.SharingItem.RemotePath.md)
- [RemoteStatus](Outlook.SharingItem.RemoteStatus.md)
- [ReplyRecipientNames](Outlook.SharingItem.ReplyRecipientNames.md)
- [ReplyRecipients](Outlook.SharingItem.ReplyRecipients.md)
- [RequestedFolder](Outlook.SharingItem.RequestedFolder.md)
- [RetentionExpirationDate](Outlook.SharingItem.RetentionExpirationDate.md)
- [RetentionPolicyName](Outlook.SharingItem.RetentionPolicyName.md)
- [RTFBody](Outlook.SharingItem.RTFBody.md)
- [Saved](Outlook.SharingItem.Saved.md)
- [SaveSentMessageFolder](Outlook.SharingItem.SaveSentMessageFolder.md)
- [SenderEmailAddress](Outlook.SharingItem.SenderEmailAddress.md)
- [SenderEmailType](Outlook.SharingItem.SenderEmailType.md)
- [SenderName](Outlook.SharingItem.SenderName.md)
- [SendUsingAccount](Outlook.SharingItem.SendUsingAccount.md)
- [Sensitivity](Outlook.SharingItem.Sensitivity.md)
- [Sent](Outlook.SharingItem.Sent.md)
- [SentOn](Outlook.SharingItem.SentOn.md)
- [SentOnBehalfOfName](Outlook.SharingItem.SentOnBehalfOfName.md)
- [Session](Outlook.SharingItem.Session.md)
- [SharingProvider](Outlook.SharingItem.SharingProvider.md)
- [SharingProviderGuid](Outlook.SharingItem.SharingProviderGuid.md)
- [Size](Outlook.SharingItem.Size.md)
- [Subject](Outlook.SharingItem.Subject.md)
- [Submitted](Outlook.SharingItem.Submitted.md)
- [TaskCompletedDate](Outlook.SharingItem.TaskCompletedDate.md)
- [TaskDueDate](Outlook.SharingItem.TaskDueDate.md)
- [TaskStartDate](Outlook.SharingItem.TaskStartDate.md)
- [TaskSubject](Outlook.SharingItem.TaskSubject.md)
- [To](Outlook.SharingItem.To.md)
- [ToDoTaskOrdinal](Outlook.SharingItem.ToDoTaskOrdinal.md)
- [Type](Outlook.SharingItem.Type.md)
- [UnRead](Outlook.SharingItem.UnRead.md)
- [UserProperties](Outlook.SharingItem.UserProperties.md)

## See also

- [Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
