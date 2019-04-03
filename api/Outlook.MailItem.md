---
title: MailItem object (Outlook)
keywords: vbaol11.chm2987
f1_keywords:
- vbaol11.chm2987
ms.prod: outlook
api_name:
- Outlook.MailItem
ms.assetid: 14197346-05d2-0250-fa4c-4a6b07daf25f
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem object (Outlook)

Represents a mail message.


## Remarks

Use the  **[CreateItem](Outlook.Application.CreateItem.md)** method to create a **MailItem** object that represents a new mail message.

Use the  **[Folder.Items](Outlook.Folder.Items.md)** property to obtain an **[Items](Outlook.Items.md)** collection representing the mail items in a folder, and the **[Items.Item](Outlook.Items.Item.md)** (_index_) method, where _index_ is the index number of a mail message or a value used to match the default property of a message, to return a single **MailItem** object from the specified folder.


## Example

The following Visual Basic for Applications (VBA) example creates and displays a new mail message.


```vb
Sub CreateMail() 
 
 Dim myItem As Object 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 myItem.Subject = "Mail to myself" 
 
 myItem.Display 
 
End Sub
```

The following VBA example sets the current folder as the Inbox and displays the second mail message in the folder. In general, the order of mail messages in a folder is not guaranteed to be in a particular order. 




```vb
Sub DisplayMail() 
 
 Dim myItem As Object 
 
 Dim myFolder As Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNamespace.GetDefaultFolder(olFolderInbox) 
 
 myFolder.Display 
 
 Set myItem = myFolder.Items(2) 
 
 myItem.Display 
 
End Sub
```


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.MailItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.MailItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.MailItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.MailItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.MailItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.MailItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.MailItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.MailItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.MailItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.MailItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.MailItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.MailItem.BeforeDelete.md)|
|[BeforeRead](Outlook.MailItem.BeforeRead.md)|
|[Close](Outlook.MailItem.Close(even).md)|
|[CustomAction](Outlook.MailItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.MailItem.CustomPropertyChange.md)|
|[Forward](Outlook.MailItem.Forward(even).md)|
|[Open](Outlook.MailItem.Open.md)|
|[PropertyChange](Outlook.MailItem.PropertyChange.md)|
|[Read](Outlook.MailItem.Read.md)|
|[ReadComplete](Outlook.mailitem.readcomplete.md)|
|[Reply](Outlook.MailItem.Reply(even).md)|
|[ReplyAll](Outlook.MailItem.ReplyAll(even).md)|
|[Send](Outlook.MailItem.Send(even).md)|
|[Unload](Outlook.MailItem.Unload.md)|
|[Write](Outlook.MailItem.Write.md)|

## Methods



|Name|
|:-----|
|[AddBusinessCard](Outlook.MailItem.AddBusinessCard.md)|
|[ClearConversationIndex](Outlook.MailItem.ClearConversationIndex.md)|
|[ClearTaskFlag](Outlook.MailItem.ClearTaskFlag.md)|
|[Close](Outlook.MailItem.Close(method).md)|
|[Copy](Outlook.MailItem.Copy.md)|
|[Delete](Outlook.MailItem.Delete.md)|
|[Display](Outlook.MailItem.Display.md)|
|[Forward](Outlook.MailItem.Forward(method).md)|
|[GetConversation](Outlook.MailItem.GetConversation.md)|
|[MarkAsTask](Outlook.MailItem.MarkAsTask.md)|
|[Move](Outlook.MailItem.Move.md)|
|[PrintOut](Outlook.MailItem.PrintOut.md)|
|[Reply](Outlook.MailItem.Reply(method).md)|
|[ReplyAll](Outlook.MailItem.ReplyAll(method).md)|
|[Save](Outlook.MailItem.Save.md)|
|[SaveAs](Outlook.MailItem.SaveAs.md)|
|[Send](Outlook.MailItem.Send(method).md)|
|[ShowCategoriesDialog](Outlook.MailItem.ShowCategoriesDialog.md)|

## Properties



|Name|
|:-----|
|[Actions](Outlook.MailItem.Actions.md)|
|[AlternateRecipientAllowed](Outlook.MailItem.AlternateRecipientAllowed.md)|
|[Application](Outlook.MailItem.Application.md)|
|[Attachments](Outlook.MailItem.Attachments.md)|
|[AutoForwarded](Outlook.MailItem.AutoForwarded.md)|
|[AutoResolvedWinner](Outlook.MailItem.AutoResolvedWinner.md)|
|[BCC](Outlook.MailItem.BCC.md)|
|[BillingInformation](Outlook.MailItem.BillingInformation.md)|
|[Body](Outlook.MailItem.Body.md)|
|[BodyFormat](Outlook.MailItem.BodyFormat.md)|
|[Categories](Outlook.MailItem.Categories.md)|
|[CC](Outlook.MailItem.CC.md)|
|[Class](Outlook.MailItem.Class.md)|
|[Companies](Outlook.MailItem.Companies.md)|
|[Conflicts](Outlook.MailItem.Conflicts.md)|
|[ConversationID](Outlook.MailItem.ConversationID.md)|
|[ConversationIndex](Outlook.MailItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.MailItem.ConversationTopic.md)|
|[CreationTime](Outlook.MailItem.CreationTime.md)|
|[DeferredDeliveryTime](Outlook.MailItem.DeferredDeliveryTime.md)|
|[DeleteAfterSubmit](Outlook.MailItem.DeleteAfterSubmit.md)|
|[DownloadState](Outlook.MailItem.DownloadState.md)|
|[EntryID](Outlook.MailItem.EntryID.md)|
|[ExpiryTime](Outlook.MailItem.ExpiryTime.md)|
|[FlagRequest](Outlook.MailItem.FlagRequest.md)|
|[FormDescription](Outlook.MailItem.FormDescription.md)|
|[GetInspector](Outlook.MailItem.GetInspector.md)|
|[HTMLBody](Outlook.MailItem.HTMLBody.md)|
|[Importance](Outlook.MailItem.Importance.md)|
|[InternetCodepage](Outlook.MailItem.InternetCodepage.md)|
|[IsConflict](Outlook.MailItem.IsConflict.md)|
|[IsMarkedAsTask](Outlook.MailItem.IsMarkedAsTask.md)|
|[ItemProperties](Outlook.MailItem.ItemProperties.md)|
|[LastModificationTime](Outlook.MailItem.LastModificationTime.md)|
|[MarkForDownload](Outlook.MailItem.MarkForDownload.md)|
|[MessageClass](Outlook.MailItem.MessageClass.md)|
|[Mileage](Outlook.MailItem.Mileage.md)|
|[NoAging](Outlook.MailItem.NoAging.md)|
|[OriginatorDeliveryReportRequested](Outlook.MailItem.OriginatorDeliveryReportRequested.md)|
|[OutlookInternalVersion](Outlook.MailItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.MailItem.OutlookVersion.md)|
|[Parent](Outlook.MailItem.Parent.md)|
|[Permission](Outlook.MailItem.Permission.md)|
|[PermissionService](Outlook.MailItem.PermissionService.md)|
|[PermissionTemplateGuid](Outlook.MailItem.PermissionTemplateGuid.md)|
|[PropertyAccessor](Outlook.MailItem.PropertyAccessor.md)|
|[ReadReceiptRequested](Outlook.MailItem.ReadReceiptRequested.md)|
|[ReceivedByEntryID](Outlook.MailItem.ReceivedByEntryID.md)|
|[ReceivedByName](Outlook.MailItem.ReceivedByName.md)|
|[ReceivedOnBehalfOfEntryID](Outlook.MailItem.ReceivedOnBehalfOfEntryID.md)|
|[ReceivedOnBehalfOfName](Outlook.MailItem.ReceivedOnBehalfOfName.md)|
|[ReceivedTime](Outlook.MailItem.ReceivedTime.md)|
|[RecipientReassignmentProhibited](Outlook.MailItem.RecipientReassignmentProhibited.md)|
|[Recipients](Outlook.MailItem.Recipients.md)|
|[ReminderOverrideDefault](Outlook.MailItem.ReminderOverrideDefault.md)|
|[ReminderPlaySound](Outlook.MailItem.ReminderPlaySound.md)|
|[ReminderSet](Outlook.MailItem.ReminderSet.md)|
|[ReminderSoundFile](Outlook.MailItem.ReminderSoundFile.md)|
|[ReminderTime](Outlook.MailItem.ReminderTime.md)|
|[RemoteStatus](Outlook.MailItem.RemoteStatus.md)|
|[ReplyRecipientNames](Outlook.MailItem.ReplyRecipientNames.md)|
|[ReplyRecipients](Outlook.MailItem.ReplyRecipients.md)|
|[RetentionExpirationDate](Outlook.MailItem.RetentionExpirationDate.md)|
|[RetentionPolicyName](Outlook.MailItem.RetentionPolicyName.md)|
|[RTFBody](Outlook.MailItem.RTFBody.md)|
|[Saved](Outlook.MailItem.Saved.md)|
|[SaveSentMessageFolder](Outlook.MailItem.SaveSentMessageFolder.md)|
|[Sender](Outlook.MailItem.Sender.md)|
|[SenderEmailAddress](Outlook.MailItem.SenderEmailAddress.md)|
|[SenderEmailType](Outlook.MailItem.SenderEmailType.md)|
|[SenderName](Outlook.MailItem.SenderName.md)|
|[SendUsingAccount](Outlook.MailItem.SendUsingAccount.md)|
|[Sensitivity](Outlook.MailItem.Sensitivity.md)|
|[Sent](Outlook.MailItem.Sent.md)|
|[SentOn](Outlook.MailItem.SentOn.md)|
|[SentOnBehalfOfName](Outlook.MailItem.SentOnBehalfOfName.md)|
|[Session](Outlook.MailItem.Session.md)|
|[Size](Outlook.MailItem.Size.md)|
|[Subject](Outlook.MailItem.Subject.md)|
|[Submitted](Outlook.MailItem.Submitted.md)|
|[TaskCompletedDate](Outlook.MailItem.TaskCompletedDate.md)|
|[TaskDueDate](Outlook.MailItem.TaskDueDate.md)|
|[TaskStartDate](Outlook.MailItem.TaskStartDate.md)|
|[TaskSubject](Outlook.MailItem.TaskSubject.md)|
|[To](Outlook.MailItem.To.md)|
|[ToDoTaskOrdinal](Outlook.MailItem.ToDoTaskOrdinal.md)|
|[UnRead](Outlook.MailItem.UnRead.md)|
|[UserProperties](Outlook.MailItem.UserProperties.md)|
|[VotingOptions](Outlook.MailItem.VotingOptions.md)|
|[VotingResponse](Outlook.MailItem.VotingResponse.md)|

## See also

- [Send an email given the SMTP address of an account](../outlook/How-to/Items-Folders-and-Stores/send-an-e-mail-given-the-smtp-address-of-an-account-outlook.md)
- [Outlook Object Model reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
