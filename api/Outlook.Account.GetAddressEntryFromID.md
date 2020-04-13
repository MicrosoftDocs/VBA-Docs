---
title: Account.GetAddressEntryFromID method (Outlook)
keywords: vbaol11.chm3427
f1_keywords:
- vbaol11.chm3427
ms.prod: outlook
api_name:
- Outlook.Account.GetAddressEntryFromID
ms.assetid: 5aa9c67e-579f-5519-ed38-c80009cf506b
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.GetAddressEntryFromID method (Outlook)

Returns an **[AddressEntry](Outlook.AddressEntry.md)** object that represents the address entry specified by the given entry ID.


## Syntax

_expression_. `GetAddressEntryFromID`( `_ID_` )

_expression_ A variable that represents an '[Account](Outlook.Account.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ID_|Required| **String**|Used to identify an address entry that is maintained for the session.|

## Return value

An **AddressEntry** that has the **[ID](Outlook.AddressEntry.ID.md)** property that matches the specified _ID_.


## Remarks

This method is similar to the  **[GetAddressEntryFromID](Outlook.NameSpace.GetAddressEntryFromID.md)** method of the **[NameSpace](Outlook.NameSpace.md)** object, but has some additional contextual information about which account to use for the look-up. If there are multiple Microsoft Exchange accounts in the current profile, use the **GetAddressEntryFromID** method for the corresponding account.

The **ID** property for an **AddressEntry** is a permanent, unique string identifier that the transport provider assigns when an **AddressEntry** is created. Outlook maintains a hierarchy of address books for a session, and the address entry that is returned must match the given ID and be in one of the address books.

 **GetAddressEntryFromID** returns an error if no item with the given ID can be found, if no connection is available, or if the user is set to work offline.


## See also


[Account Object](Outlook.Account.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]