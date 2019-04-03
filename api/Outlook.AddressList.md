---
title: AddressList object (Outlook)
keywords: vbaol11.chm2022
f1_keywords:
- vbaol11.chm2022
ms.prod: outlook
api_name:
- Outlook.AddressList
ms.assetid: 84611afe-48b1-185b-df4b-0f004e7436ff
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressList object (Outlook)

Represents an address book that contains a set of  **[AddressEntry](Outlook.AddressEntry.md)** objects.


## Remarks

The  **AddressList** object is an address book that contains a set of **[AddressEntry](Outlook.AddressEntry.md)** objects.

The  **AddressList** object supplies a list of address entries to which a messaging system can deliver messages. An **AddressList** object represents one address book container available under the transport provider's address book hierarchy for the current session. The entire hierarchy is available through the parent **[AddressLists](Outlook.AddressLists.md)** object.


## Example

The following example retrieves an  **AddressList** object that represents the Personal Address List.


```vb
Set myAddressList = Application.Session.AddressLists("Personal Address Book")
```


## Methods



|Name|
|:-----|
|[GetContactsFolder](Outlook.AddressList.GetContactsFolder.md)|

## Properties



|Name|
|:-----|
|[AddressEntries](Outlook.AddressList.AddressEntries.md)|
|[AddressListType](Outlook.AddressList.AddressListType.md)|
|[Application](Outlook.AddressList.Application.md)|
|[Class](Outlook.AddressList.Class.md)|
|[ID](Outlook.AddressList.ID.md)|
|[Index](Outlook.AddressList.Index.md)|
|[IsInitialAddressList](Outlook.AddressList.IsInitialAddressList.md)|
|[IsReadOnly](Outlook.AddressList.IsReadOnly.md)|
|[Name](Outlook.AddressList.Name.md)|
|[Parent](Outlook.AddressList.Parent.md)|
|[PropertyAccessor](Outlook.AddressList.PropertyAccessor.md)|
|[ResolutionOrder](Outlook.AddressList.ResolutionOrder.md)|
|[Session](Outlook.AddressList.Session.md)|

## See also


[AddressList Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
