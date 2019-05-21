---
title: NameSpace.GetStoreFromID method (Outlook)
keywords: vbaol11.chm786
f1_keywords:
- vbaol11.chm786
ms.prod: outlook
api_name:
- Outlook.NameSpace.GetStoreFromID
ms.assetid: ba5b3df8-22a5-39fa-68ab-9f1e4cfe7f47
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.GetStoreFromID method (Outlook)

Returns a  **[Store](Outlook.Store.md)** object that represents the store specified by _ID_.


## Syntax

_expression_. `GetStoreFromID`( `_ID_` )

_expression_ A variable that represents a [NameSpace](Outlook.NameSpace.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ID_|Required| **String**|A string value identifying a store.|

## Return value

A  **Store** object that has the **[StoreID](Outlook.Store.StoreID.md)** property matching _ID_.


## Remarks

The  **StoreID** property of a **Store** is unique to the profile for the session. It is equivalent to the MAPI property **PR_STORE_ENTRY_ID**.

The store must be mounted in order for this method to succeed.

 **GetStoreFromID** returns an error if no store with the specified _ID_ can be found for the current session.


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]