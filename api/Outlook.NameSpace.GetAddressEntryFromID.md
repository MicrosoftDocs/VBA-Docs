---
title: NameSpace.GetAddressEntryFromID method (Outlook)
keywords: vbaol11.chm784
f1_keywords:
- vbaol11.chm784
ms.prod: outlook
api_name:
- Outlook.NameSpace.GetAddressEntryFromID
ms.assetid: 04e9d2c5-231d-35c8-eafa-0e58fbd7a2a1
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.GetAddressEntryFromID method (Outlook)

Returns an **[AddressEntry](Outlook.AddressEntry.md)** object that represents the address entry for the specified _ID_.


## Syntax

_expression_. `GetAddressEntryFromID`( `_ID_` )

_expression_ A variable that represents a '[NameSpace](Outlook.NameSpace.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ID_|Required| **String**|Used to identify an address entry that is maintained for the session.|

## Return value

An **AddressEntry** that has the **[ID](Outlook.AddressEntry.ID.md)** property that matches the specified _ID_.


## Remarks

This method is similar to the  **[GetAddressEntryFromID](Outlook.Account.GetAddressEntryFromID.md)** method of the **[Account](Outlook.Account.md)** object. Use this method if there is only the primary Exchange account in the current profile. If there are multiple Microsoft Exchange accounts in the current profile, use the **GetAddressEntryFromID** method for the corresponding account.

The **ID** property for an **AddressEntry** is a permanent, unique string identifier that the transport provider assigns when an **AddressEntry** is created.

Outlook maintains a hierarchy of address books for a session, and the address entry returned must match the given  _ID_ and be in one of the address books.

 **GetAddressEntryFromID** returns an error if no item with the given _ID_ can be found, if no connection is available, or if the user is set to work offline.


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]