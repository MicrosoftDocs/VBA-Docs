---
title: AddressList Object (Outlook)
keywords: vbaol11.chm2022
f1_keywords:
- vbaol11.chm2022
ms.prod: outlook
api_name:
- Outlook.AddressList
ms.assetid: 84611afe-48b1-185b-df4b-0f004e7436ff
ms.date: 06/08/2017
---


# AddressList Object (Outlook)

Represents an address book that contains a set of  **[AddressEntry](../../../api/Outlook.AddressEntry.md)** objects.


## Remarks

The  **AddressList** object is an address book that contains a set of **[AddressEntry](../../../api/Outlook.AddressEntry.md)** objects.

The  **AddressList** object supplies a list of address entries to which a messaging system can deliver messages. An **AddressList** object represents one address book container available under the transport provider's address book hierarchy for the current session. The entire hierarchy is available through the parent **[AddressLists](../../../api/Outlook.AddressLists.md)** object.


## Example

The following example retrieves an  **AddressList** object that represents the Personal Address List.


```
Set myAddressList = Application.Session.AddressLists("Personal Address Book")
```


## Methods



|**Name**|
|:-----|
|[GetContactsFolder](../../../api/Outlook.AddressList.GetContactsFolder.md)|

## Properties



|**Name**|
|:-----|
|[AddressEntries](../../../api/Outlook.AddressList.AddressEntries.md)|
|[AddressListType](../../../api/Outlook.AddressList.AddressListType.md)|
|[Application](../../../api/Outlook.AddressList.Application.md)|
|[Class](../../../api/Outlook.AddressList.Class.md)|
|[ID](../../../api/Outlook.AddressList.ID.md)|
|[Index](../../../api/Outlook.AddressList.Index.md)|
|[IsInitialAddressList](../../../api/Outlook.AddressList.IsInitialAddressList.md)|
|[IsReadOnly](../../../api/Outlook.AddressList.IsReadOnly.md)|
|[Name](../../../api/Outlook.AddressList.Name.md)|
|[Parent](../../../api/Outlook.AddressList.Parent.md)|
|[PropertyAccessor](../../../api/Outlook.AddressList.PropertyAccessor.md)|
|[ResolutionOrder](../../../api/Outlook.AddressList.ResolutionOrder.md)|
|[Session](../../../api/Outlook.AddressList.Session.md)|

## See also


#### Other resources


[AddressList Object Members](../../../api/overview/Outlook.md)
[Outlook Object Model Reference](../../../api/overview/object-model-outlook-vba-reference.md)
