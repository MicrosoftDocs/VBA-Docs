---
title: ExchangeUser object (Outlook)
keywords: vbaol11.chm3158
f1_keywords:
- vbaol11.chm3158
ms.prod: outlook
api_name:
- Outlook.ExchangeUser
ms.assetid: 6ec117d1-7fdb-aa36-b567-1242f8238df0
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser object (Outlook)

Provides detailed information about an **[AddressEntry](Outlook.AddressEntry.md)** that represents a Microsoft Exchange mailbox user.


## Remarks

 **ExchangeUser** is derived from the **AddressEntry** object, and is returned instead of an **AddressEntry** when the caller performs a query interface on the **AddressEntry** object.

This object provides first-class access to properties applicable to Exchange users such as  **[FirstName](Outlook.ExchangeUser.FirstName.md)**, **[JobTitle](Outlook.ExchangeUser.JobTitle.md)**, **[LastName](Outlook.ExchangeUser.LastName.md)**, and **[OfficeLocation](Outlook.ExchangeUser.OfficeLocation.md)**. You can also access other properties specific to the Exchange user that are not exposed in the object model through the **[PropertyAccessor](Outlook.PropertyAccessor.md)** object. Note that some of the explicit built-in properties are read-write properties. Setting these properties requires the code to be running under an appropriate Exchange administrator account; without sufficient permissions, calling the **[ExchangeUser.Update](Outlook.ExchangeUser.Update.md)** method will result in a "permission denied" error.


## Example

The following code sample shows how to obtain the business phone number, office location, and job title for all entries in the Exchange Global Address List.


```vb
Sub DemoAE() 
 
 Dim colAL As Outlook.AddressLists 
 
 Dim oAL As Outlook.AddressList 
 
 Dim colAE As Outlook.AddressEntries 
 
 Dim oAE As Outlook.AddressEntry 
 
 Dim oExUser As Outlook.ExchangeUser 
 
 Set colAL = Application.Session.AddressLists 
 
 For Each oAL In colAL 
 
 'Address list is an Exchange Global Address List 
 
 If oAL.AddressListType = olExchangeGlobalAddressList Then 
 
 Set colAE = oAL.AddressEntries 
 
 For Each oAE In colAE 
 
 If oAE.AddressEntryUserType = _ 
 
 olExchangeUserAddressEntry Then 
 
 Set oExUser = oAE.GetExchangeUser 
 
 Debug.Print(oExUser.JobTitle) 
 
 Debug.Print(oExUser.OfficeLocation) 
 
 Debug.Print(oExUser.BusinessTelephoneNumber) 
 
 End If 
 
 Next 
 
 End If 
 
 Next 
 
End Sub
```


## Methods



|Name|
|:-----|
|[Delete](Outlook.ExchangeUser.Delete.md)|
|[Details](Outlook.ExchangeUser.Details.md)|
|[GetContact](Outlook.ExchangeUser.GetContact.md)|
|[GetDirectReports](Outlook.ExchangeUser.GetDirectReports.md)|
|[GetExchangeDistributionList](Outlook.ExchangeUser.GetExchangeDistributionList.md)|
|[GetExchangeUser](Outlook.ExchangeUser.GetExchangeUser.md)|
|[GetExchangeUserManager](Outlook.ExchangeUser.GetExchangeUserManager.md)|
|[GetFreeBusy](Outlook.ExchangeUser.GetFreeBusy.md)|
|[GetMemberOfList](Outlook.ExchangeUser.GetMemberOfList.md)|
|[GetPicture](Outlook.ExchangeUser.GetPicture.md)|
|[Update](Outlook.ExchangeUser.Update.md)|
|[GetUnifiedGroup](Outlook.exchangeuser.getunifiedgroup.md)|
|[GetUnifiedGroupFromStore](Outlook.exchangeuser.getunifiedgroupfromstore.md)|
|[IsUnifiedGroup](Outlook.exchangeuser.isunifiedgroup.md)|

## Properties



|Name|
|:-----|
|[Address](Outlook.ExchangeUser.Address.md)|
|[AddressEntryUserType](Outlook.ExchangeUser.AddressEntryUserType.md)|
|[Alias](Outlook.ExchangeUser.Alias.md)|
|[Application](Outlook.ExchangeUser.Application.md)|
|[AssistantName](Outlook.ExchangeUser.AssistantName.md)|
|[BusinessTelephoneNumber](Outlook.ExchangeUser.BusinessTelephoneNumber.md)|
|[City](Outlook.ExchangeUser.City.md)|
|[Class](Outlook.ExchangeUser.Class.md)|
|[Comments](Outlook.ExchangeUser.Comments.md)|
|[CompanyName](Outlook.ExchangeUser.CompanyName.md)|
|[Department](Outlook.ExchangeUser.Department.md)|
|[DisplayType](Outlook.ExchangeUser.DisplayType.md)|
|[FirstName](Outlook.ExchangeUser.FirstName.md)|
|[ID](Outlook.ExchangeUser.ID.md)|
|[JobTitle](Outlook.ExchangeUser.JobTitle.md)|
|[LastName](Outlook.ExchangeUser.LastName.md)|
|[MobileTelephoneNumber](Outlook.ExchangeUser.MobileTelephoneNumber.md)|
|[Name](Outlook.ExchangeUser.Name.md)|
|[OfficeLocation](Outlook.ExchangeUser.OfficeLocation.md)|
|[Parent](Outlook.ExchangeUser.Parent.md)|
|[PostalCode](Outlook.ExchangeUser.PostalCode.md)|
|[PrimarySmtpAddress](Outlook.ExchangeUser.PrimarySmtpAddress.md)|
|[PropertyAccessor](Outlook.ExchangeUser.PropertyAccessor.md)|
|[Session](Outlook.ExchangeUser.Session.md)|
|[StateOrProvince](Outlook.ExchangeUser.StateOrProvince.md)|
|[StreetAddress](Outlook.ExchangeUser.StreetAddress.md)|
|[Type](Outlook.ExchangeUser.Type.md)|
|[YomiCompanyName](Outlook.ExchangeUser.YomiCompanyName.md)|
|[YomiDepartment](Outlook.ExchangeUser.YomiDepartment.md)|
|[YomiDisplayName](Outlook.ExchangeUser.YomiDisplayName.md)|
|[YomiFirstName](Outlook.ExchangeUser.YomiFirstName.md)|
|[YomiLastName](Outlook.ExchangeUser.YomiLastName.md)|

## See also


[ExchangeUser Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
