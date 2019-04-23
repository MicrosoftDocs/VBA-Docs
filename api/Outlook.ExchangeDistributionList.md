---
title: ExchangeDistributionList object (Outlook)
keywords: vbaol11.chm3159
f1_keywords:
- vbaol11.chm3159
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList
ms.assetid: 2830dfba-6c0a-a81f-6b98-92ac2aafb59d
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeDistributionList object (Outlook)

The  **ExchangeDistributionList** object provides detailed information about an **[AddressEntry](Outlook.AddressEntry.md)** that represents an Exchange distribution list.


## Remarks

 **ExchangeDistributionList** is a derived class of **AddressEntry**, and is returned instead of an **AddressEntry** when the caller performs a **QueryInterface** on the **AddressEntry**.

The  **AddressEntry.Members** property supports enumerating members of a distribution list. **ExchangeDistributionList** adds the first-class properties for **[Alias](Outlook.ExchangeDistributionList.Alias.md)**, **[Comments](Outlook.ExchangeDistributionList.Comments.md)**, and **[PrimarySmtpAddress](Outlook.ExchangeDistributionList.PrimarySmtpAddress.md)**. You can also access other properties specific to the Exchange distribution list that are not exposed in the object model through the **[PropertyAccessor](Outlook.PropertyAccessor.md)** object.

Some properties such as  **Comments** are read-write properties. Setting these properties requires the code to be running under an appropriate Exchange administrator account; without sufficient permissions, calling the **[ExchangeUser.Update](Outlook.ExchangeUser.Update.md)** method will result in a "permission denied" error.


## Example

The following code sample shows how to obtain the names of the Exchange distribution lists that the current user's manager belongs to. It uses the  **[ExchangeUser.GetExchangeUserManager](Outlook.ExchangeUser.GetExchangeUserManager.md)** method to obtain information about the user's manager, and uses **[ExchangeUser.GetMemberOfList](Outlook.ExchangeUser.GetMemberOfList.md)** to obtain the distribution lists (represented by **ExchangeDistributionList** objects) that the manager has joined.


```vb
Sub ShowManagerDistLists() 
 Dim oAE As Outlook.AddressEntry 
 Dim oExUser As Outlook.ExchangeUser 
 Dim oDistListEntries As Outlook.AddressEntries 
 
 'Obtain the AddressEntry for CurrentUser 
 Set oExUser = _ 
 Application.Session.CurrentUser.AddressEntry.GetExchangeUser 
 
 'Obtain distribution lists that the user's manager has joined 
 Set oDistListEntries = oExUser.GetExchangeUserManager.GetMemberOfList 
 For Each oAE In oDistListEntries 
 If oAE.AddressEntryUserType = _ 
 olExchangeDistributionListAddressEntry Then 
 Debug.Print (oAE.name) 
 End If 
 Next 
End Sub 

```


## Methods



|Name|
|:-----|
|[Delete](Outlook.ExchangeDistributionList.Delete.md)|
|[Details](Outlook.ExchangeDistributionList.Details.md)|
|[GetContact](Outlook.ExchangeDistributionList.GetContact.md)|
|[GetExchangeDistributionList](Outlook.ExchangeDistributionList.GetExchangeDistributionList.md)|
|[GetExchangeDistributionListMembers](Outlook.ExchangeDistributionList.GetExchangeDistributionListMembers.md)|
|[GetExchangeUser](Outlook.ExchangeDistributionList.GetExchangeUser.md)|
|[GetFreeBusy](Outlook.ExchangeDistributionList.GetFreeBusy.md)|
|[GetMemberOfList](Outlook.ExchangeDistributionList.GetMemberOfList.md)|
|[GetOwners](Outlook.ExchangeDistributionList.GetOwners.md)|
|[Update](Outlook.ExchangeDistributionList.Update.md)|
|[GetUnifiedGroup](Outlook.exchangedistributionlist.getunifiedgroup.md)|
|[GetUnifiedGroupFromStore](Outlook.exchangedistributionlist.getunifiedgroupfromstore.md)|
|[IsUnifiedGroup](Outlook.exchangedistributionlist.isunifiedgroup.md)|

## Properties



|Name|
|:-----|
|[Address](Outlook.ExchangeDistributionList.Address.md)|
|[AddressEntryUserType](Outlook.ExchangeDistributionList.AddressEntryUserType.md)|
|[Alias](Outlook.ExchangeDistributionList.Alias.md)|
|[Application](Outlook.ExchangeDistributionList.Application.md)|
|[Class](Outlook.ExchangeDistributionList.Class.md)|
|[Comments](Outlook.ExchangeDistributionList.Comments.md)|
|[DisplayType](Outlook.ExchangeDistributionList.DisplayType.md)|
|[ID](Outlook.ExchangeDistributionList.ID.md)|
|[Name](Outlook.ExchangeDistributionList.Name.md)|
|[Parent](Outlook.ExchangeDistributionList.Parent.md)|
|[PrimarySmtpAddress](Outlook.ExchangeDistributionList.PrimarySmtpAddress.md)|
|[PropertyAccessor](Outlook.ExchangeDistributionList.PropertyAccessor.md)|
|[Session](Outlook.ExchangeDistributionList.Session.md)|
|[Type](Outlook.ExchangeDistributionList.Type.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]