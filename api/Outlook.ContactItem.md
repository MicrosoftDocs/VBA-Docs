---
title: ContactItem object (Outlook)
keywords: vbaol11.chm2992
f1_keywords:
- vbaol11.chm2992
ms.prod: outlook
api_name:
- Outlook.ContactItem
ms.assetid: 8e32093c-a678-f1fd-3f35-c2d8994d166f
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem object (Outlook)

Represents a contact in a Contacts folder.


## Remarks

A contact can represent any person with whom you have any personal or professional contact.

Use the  **[CreateItem](Outlook.Application.CreateItem.md)** method to create a **ContactItem** object that represents a new contact.

Use  **[Items](Outlook.Folder.Items.md)** (_index_), where _index_ is the index number of a contact or a value used to match the default property of a contact, to return a single **ContactItem** object from a Contacts folder.


## Example

The following Visual Basic for Applications (VBA) example returns a new contact.


```vb
Set myItem = Application.CreateItem(olContactItem)
```


## Events



|Name|
|:-----|
|[AfterWrite](Outlook.ContactItem.AfterWrite.md)|
|[AttachmentAdd](Outlook.ContactItem.AttachmentAdd.md)|
|[AttachmentRead](Outlook.ContactItem.AttachmentRead.md)|
|[AttachmentRemove](Outlook.ContactItem.AttachmentRemove.md)|
|[BeforeAttachmentAdd](Outlook.ContactItem.BeforeAttachmentAdd.md)|
|[BeforeAttachmentPreview](Outlook.ContactItem.BeforeAttachmentPreview.md)|
|[BeforeAttachmentRead](Outlook.ContactItem.BeforeAttachmentRead.md)|
|[BeforeAttachmentSave](Outlook.ContactItem.BeforeAttachmentSave.md)|
|[BeforeAttachmentWriteToTempFile](Outlook.ContactItem.BeforeAttachmentWriteToTempFile.md)|
|[BeforeAutoSave](Outlook.ContactItem.BeforeAutoSave.md)|
|[BeforeCheckNames](Outlook.ContactItem.BeforeCheckNames.md)|
|[BeforeDelete](Outlook.ContactItem.BeforeDelete.md)|
|[BeforeRead](Outlook.ContactItem.BeforeRead.md)|
|[Close](Outlook.ContactItem.Close(even).md)|
|[CustomAction](Outlook.ContactItem.CustomAction.md)|
|[CustomPropertyChange](Outlook.ContactItem.CustomPropertyChange.md)|
|[Forward](Outlook.ContactItem.Forward.md)|
|[Open](Outlook.ContactItem.Open.md)|
|[PropertyChange](Outlook.ContactItem.PropertyChange.md)|
|[Read](Outlook.ContactItem.Read.md)|
|[ReadComplete](Outlook.contactitem.readcomplete.md)|
|[Reply](Outlook.ContactItem.Reply.md)|
|[ReplyAll](Outlook.ContactItem.ReplyAll.md)|
|[Send](Outlook.ContactItem.Send.md)|
|[Unload](Outlook.ContactItem.Unload.md)|
|[Write](Outlook.ContactItem.Write.md)|

## Methods



|Name|
|:-----|
|[AddBusinessCardLogoPicture](Outlook.ContactItem.AddBusinessCardLogoPicture.md)|
|[AddPicture](Outlook.ContactItem.AddPicture.md)|
|[ClearTaskFlag](Outlook.ContactItem.ClearTaskFlag.md)|
|[Close](Outlook.ContactItem.Close(method).md)|
|[Copy](Outlook.ContactItem.Copy.md)|
|[Delete](Outlook.ContactItem.Delete.md)|
|[Display](Outlook.ContactItem.Display.md)|
|[ForwardAsBusinessCard](Outlook.ContactItem.ForwardAsBusinessCard.md)|
|[ForwardAsVcard](Outlook.ContactItem.ForwardAsVcard.md)|
|[GetConversation](Outlook.ContactItem.GetConversation.md)|
|[MarkAsTask](Outlook.ContactItem.MarkAsTask.md)|
|[Move](Outlook.ContactItem.Move.md)|
|[PrintOut](Outlook.ContactItem.PrintOut.md)|
|[RemovePicture](Outlook.ContactItem.RemovePicture.md)|
|[ResetBusinessCard](Outlook.ContactItem.ResetBusinessCard.md)|
|[Save](Outlook.ContactItem.Save.md)|
|[SaveAs](Outlook.ContactItem.SaveAs.md)|
|[SaveBusinessCardImage](Outlook.ContactItem.SaveBusinessCardImage.md)|
|[ShowBusinessCardEditor](Outlook.ContactItem.ShowBusinessCardEditor.md)|
|[ShowCategoriesDialog](Outlook.ContactItem.ShowCategoriesDialog.md)|
|[ShowCheckAddressDialog](Outlook.contactitem.showcheckaddressdialog.md)|
|[ShowCheckFullNameDialog](Outlook.contactitem.showcheckfullnamedialog.md)|
|[ShowCheckPhoneDialog](Outlook.ContactItem.ShowCheckPhoneDialog.md)|

## Properties



|Name|
|:-----|
|[Account](Outlook.ContactItem.Account.md)|
|[Actions](Outlook.ContactItem.Actions.md)|
|[Anniversary](Outlook.ContactItem.Anniversary.md)|
|[Application](Outlook.ContactItem.Application.md)|
|[AssistantName](Outlook.ContactItem.AssistantName.md)|
|[AssistantTelephoneNumber](Outlook.ContactItem.AssistantTelephoneNumber.md)|
|[Attachments](Outlook.ContactItem.Attachments.md)|
|[AutoResolvedWinner](Outlook.ContactItem.AutoResolvedWinner.md)|
|[BillingInformation](Outlook.ContactItem.BillingInformation.md)|
|[Birthday](Outlook.ContactItem.Birthday.md)|
|[Body](Outlook.ContactItem.Body.md)|
|[Business2TelephoneNumber](Outlook.ContactItem.Business2TelephoneNumber.md)|
|[BusinessAddress](Outlook.ContactItem.BusinessAddress.md)|
|[BusinessAddressCity](Outlook.ContactItem.BusinessAddressCity.md)|
|[BusinessAddressCountry](Outlook.ContactItem.BusinessAddressCountry.md)|
|[BusinessAddressPostalCode](Outlook.ContactItem.BusinessAddressPostalCode.md)|
|[BusinessAddressPostOfficeBox](Outlook.ContactItem.BusinessAddressPostOfficeBox.md)|
|[BusinessAddressState](Outlook.ContactItem.BusinessAddressState.md)|
|[BusinessAddressStreet](Outlook.ContactItem.BusinessAddressStreet.md)|
|[BusinessCardLayoutXml](Outlook.ContactItem.BusinessCardLayoutXml.md)|
|[BusinessCardType](Outlook.ContactItem.BusinessCardType.md)|
|[BusinessFaxNumber](Outlook.ContactItem.BusinessFaxNumber.md)|
|[BusinessHomePage](Outlook.ContactItem.BusinessHomePage.md)|
|[BusinessTelephoneNumber](Outlook.ContactItem.BusinessTelephoneNumber.md)|
|[CallbackTelephoneNumber](Outlook.ContactItem.CallbackTelephoneNumber.md)|
|[CarTelephoneNumber](Outlook.ContactItem.CarTelephoneNumber.md)|
|[Categories](Outlook.ContactItem.Categories.md)|
|[Children](Outlook.ContactItem.Children.md)|
|[Class](Outlook.ContactItem.Class.md)|
|[Companies](Outlook.ContactItem.Companies.md)|
|[CompanyAndFullName](Outlook.ContactItem.CompanyAndFullName.md)|
|[CompanyLastFirstNoSpace](Outlook.ContactItem.CompanyLastFirstNoSpace.md)|
|[CompanyLastFirstSpaceOnly](Outlook.ContactItem.CompanyLastFirstSpaceOnly.md)|
|[CompanyMainTelephoneNumber](Outlook.ContactItem.CompanyMainTelephoneNumber.md)|
|[CompanyName](Outlook.ContactItem.CompanyName.md)|
|[ComputerNetworkName](Outlook.ContactItem.ComputerNetworkName.md)|
|[Conflicts](Outlook.ContactItem.Conflicts.md)|
|[ConversationID](Outlook.ContactItem.ConversationID.md)|
|[ConversationIndex](Outlook.ContactItem.ConversationIndex.md)|
|[ConversationTopic](Outlook.ContactItem.ConversationTopic.md)|
|[CreationTime](Outlook.ContactItem.CreationTime.md)|
|[CustomerID](Outlook.ContactItem.CustomerID.md)|
|[Department](Outlook.ContactItem.Department.md)|
|[DownloadState](Outlook.ContactItem.DownloadState.md)|
|[Email1Address](Outlook.ContactItem.Email1Address.md)|
|[Email1AddressType](Outlook.ContactItem.Email1AddressType.md)|
|[Email1DisplayName](Outlook.ContactItem.Email1DisplayName.md)|
|[Email1EntryID](Outlook.ContactItem.Email1EntryID.md)|
|[Email2Address](Outlook.ContactItem.Email2Address.md)|
|[Email2AddressType](Outlook.ContactItem.Email2AddressType.md)|
|[Email2DisplayName](Outlook.ContactItem.Email2DisplayName.md)|
|[Email2EntryID](Outlook.ContactItem.Email2EntryID.md)|
|[Email3Address](Outlook.ContactItem.Email3Address.md)|
|[Email3AddressType](Outlook.ContactItem.Email3AddressType.md)|
|[Email3DisplayName](Outlook.ContactItem.Email3DisplayName.md)|
|[Email3EntryID](Outlook.ContactItem.Email3EntryID.md)|
|[EntryID](Outlook.ContactItem.EntryID.md)|
|[FileAs](Outlook.ContactItem.FileAs.md)|
|[FirstName](Outlook.ContactItem.FirstName.md)|
|[FormDescription](Outlook.ContactItem.FormDescription.md)|
|[FTPSite](Outlook.ContactItem.FTPSite.md)|
|[FullName](Outlook.ContactItem.FullName.md)|
|[FullNameAndCompany](Outlook.ContactItem.FullNameAndCompany.md)|
|[Gender](Outlook.ContactItem.Gender.md)|
|[GetInspector](Outlook.ContactItem.GetInspector.md)|
|[GovernmentIDNumber](Outlook.ContactItem.GovernmentIDNumber.md)|
|[HasPicture](Outlook.ContactItem.HasPicture.md)|
|[Hobby](Outlook.ContactItem.Hobby.md)|
|[Home2TelephoneNumber](Outlook.ContactItem.Home2TelephoneNumber.md)|
|[HomeAddress](Outlook.ContactItem.HomeAddress.md)|
|[HomeAddressCity](Outlook.ContactItem.HomeAddressCity.md)|
|[HomeAddressCountry](Outlook.ContactItem.HomeAddressCountry.md)|
|[HomeAddressPostalCode](Outlook.ContactItem.HomeAddressPostalCode.md)|
|[HomeAddressPostOfficeBox](Outlook.ContactItem.HomeAddressPostOfficeBox.md)|
|[HomeAddressState](Outlook.ContactItem.HomeAddressState.md)|
|[HomeAddressStreet](Outlook.ContactItem.HomeAddressStreet.md)|
|[HomeFaxNumber](Outlook.ContactItem.HomeFaxNumber.md)|
|[HomeTelephoneNumber](Outlook.ContactItem.HomeTelephoneNumber.md)|
|[IMAddress](Outlook.ContactItem.IMAddress.md)|
|[Importance](Outlook.ContactItem.Importance.md)|
|[Initials](Outlook.ContactItem.Initials.md)|
|[InternetFreeBusyAddress](Outlook.ContactItem.InternetFreeBusyAddress.md)|
|[IsConflict](Outlook.ContactItem.IsConflict.md)|
|[ISDNNumber](Outlook.ContactItem.ISDNNumber.md)|
|[IsMarkedAsTask](Outlook.ContactItem.IsMarkedAsTask.md)|
|[ItemProperties](Outlook.ContactItem.ItemProperties.md)|
|[JobTitle](Outlook.ContactItem.JobTitle.md)|
|[Journal](Outlook.ContactItem.Journal.md)|
|[Language](Outlook.ContactItem.Language.md)|
|[LastFirstAndSuffix](Outlook.ContactItem.LastFirstAndSuffix.md)|
|[LastFirstNoSpace](Outlook.ContactItem.LastFirstNoSpace.md)|
|[LastFirstNoSpaceAndSuffix](Outlook.ContactItem.LastFirstNoSpaceAndSuffix.md)|
|[LastFirstNoSpaceCompany](Outlook.ContactItem.LastFirstNoSpaceCompany.md)|
|[LastFirstSpaceOnly](Outlook.ContactItem.LastFirstSpaceOnly.md)|
|[LastFirstSpaceOnlyCompany](Outlook.ContactItem.LastFirstSpaceOnlyCompany.md)|
|[LastModificationTime](Outlook.ContactItem.LastModificationTime.md)|
|[LastName](Outlook.ContactItem.LastName.md)|
|[LastNameAndFirstName](Outlook.ContactItem.LastNameAndFirstName.md)|
|[MailingAddress](Outlook.ContactItem.MailingAddress.md)|
|[MailingAddressCity](Outlook.ContactItem.MailingAddressCity.md)|
|[MailingAddressCountry](Outlook.ContactItem.MailingAddressCountry.md)|
|[MailingAddressPostalCode](Outlook.ContactItem.MailingAddressPostalCode.md)|
|[MailingAddressPostOfficeBox](Outlook.ContactItem.MailingAddressPostOfficeBox.md)|
|[MailingAddressState](Outlook.ContactItem.MailingAddressState.md)|
|[MailingAddressStreet](Outlook.ContactItem.MailingAddressStreet.md)|
|[ManagerName](Outlook.ContactItem.ManagerName.md)|
|[MarkForDownload](Outlook.ContactItem.MarkForDownload.md)|
|[MessageClass](Outlook.ContactItem.MessageClass.md)|
|[MiddleName](Outlook.ContactItem.MiddleName.md)|
|[Mileage](Outlook.ContactItem.Mileage.md)|
|[MobileTelephoneNumber](Outlook.ContactItem.MobileTelephoneNumber.md)|
|[NetMeetingAlias](Outlook.ContactItem.NetMeetingAlias.md)|
|[NetMeetingServer](Outlook.ContactItem.NetMeetingServer.md)|
|[NickName](Outlook.ContactItem.NickName.md)|
|[NoAging](Outlook.ContactItem.NoAging.md)|
|[OfficeLocation](Outlook.ContactItem.OfficeLocation.md)|
|[OrganizationalIDNumber](Outlook.ContactItem.OrganizationalIDNumber.md)|
|[OtherAddress](Outlook.ContactItem.OtherAddress.md)|
|[OtherAddressCity](Outlook.ContactItem.OtherAddressCity.md)|
|[OtherAddressCountry](Outlook.ContactItem.OtherAddressCountry.md)|
|[OtherAddressPostalCode](Outlook.ContactItem.OtherAddressPostalCode.md)|
|[OtherAddressPostOfficeBox](Outlook.ContactItem.OtherAddressPostOfficeBox.md)|
|[OtherAddressState](Outlook.ContactItem.OtherAddressState.md)|
|[OtherAddressStreet](Outlook.ContactItem.OtherAddressStreet.md)|
|[OtherFaxNumber](Outlook.ContactItem.OtherFaxNumber.md)|
|[OtherTelephoneNumber](Outlook.ContactItem.OtherTelephoneNumber.md)|
|[OutlookInternalVersion](Outlook.ContactItem.OutlookInternalVersion.md)|
|[OutlookVersion](Outlook.ContactItem.OutlookVersion.md)|
|[PagerNumber](Outlook.ContactItem.PagerNumber.md)|
|[Parent](Outlook.ContactItem.Parent.md)|
|[PersonalHomePage](Outlook.ContactItem.PersonalHomePage.md)|
|[PrimaryTelephoneNumber](Outlook.ContactItem.PrimaryTelephoneNumber.md)|
|[Profession](Outlook.ContactItem.Profession.md)|
|[PropertyAccessor](Outlook.ContactItem.PropertyAccessor.md)|
|[RadioTelephoneNumber](Outlook.ContactItem.RadioTelephoneNumber.md)|
|[ReferredBy](Outlook.ContactItem.ReferredBy.md)|
|[ReminderOverrideDefault](Outlook.ContactItem.ReminderOverrideDefault.md)|
|[ReminderPlaySound](Outlook.ContactItem.ReminderPlaySound.md)|
|[ReminderSet](Outlook.ContactItem.ReminderSet.md)|
|[ReminderSoundFile](Outlook.ContactItem.ReminderSoundFile.md)|
|[ReminderTime](Outlook.ContactItem.ReminderTime.md)|
|[RTFBody](Outlook.ContactItem.RTFBody.md)|
|[Saved](Outlook.ContactItem.Saved.md)|
|[SelectedMailingAddress](Outlook.ContactItem.SelectedMailingAddress.md)|
|[Sensitivity](Outlook.ContactItem.Sensitivity.md)|
|[Session](Outlook.ContactItem.Session.md)|
|[Size](Outlook.ContactItem.Size.md)|
|[Spouse](Outlook.ContactItem.Spouse.md)|
|[Subject](Outlook.ContactItem.Subject.md)|
|[Suffix](Outlook.ContactItem.Suffix.md)|
|[TaskCompletedDate](Outlook.ContactItem.TaskCompletedDate.md)|
|[TaskDueDate](Outlook.ContactItem.TaskDueDate.md)|
|[TaskStartDate](Outlook.ContactItem.TaskStartDate.md)|
|[TaskSubject](Outlook.ContactItem.TaskSubject.md)|
|[TelexNumber](Outlook.ContactItem.TelexNumber.md)|
|[Title](Outlook.ContactItem.Title.md)|
|[ToDoTaskOrdinal](Outlook.ContactItem.ToDoTaskOrdinal.md)|
|[TTYTDDTelephoneNumber](Outlook.ContactItem.TTYTDDTelephoneNumber.md)|
|[UnRead](Outlook.ContactItem.UnRead.md)|
|[User1](Outlook.ContactItem.User1.md)|
|[User2](Outlook.ContactItem.User2.md)|
|[User3](Outlook.ContactItem.User3.md)|
|[User4](Outlook.ContactItem.User4.md)|
|[UserProperties](Outlook.ContactItem.UserProperties.md)|
|[WebPage](Outlook.ContactItem.WebPage.md)|
|[YomiCompanyName](Outlook.ContactItem.YomiCompanyName.md)|
|[YomiFirstName](Outlook.ContactItem.YomiFirstName.md)|
|[YomiLastName](Outlook.ContactItem.YomiLastName.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[ContactItem Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
