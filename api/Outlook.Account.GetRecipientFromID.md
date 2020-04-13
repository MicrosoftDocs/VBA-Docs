---
title: Account.GetRecipientFromID method (Outlook)
keywords: vbaol11.chm3428
f1_keywords:
- vbaol11.chm3428
ms.prod: outlook
api_name:
- Outlook.Account.GetRecipientFromID
ms.assetid: 7b97ce67-6015-ece6-de1b-6d4226be83aa
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.GetRecipientFromID method (Outlook)

Returns the **[Recipient](Outlook.Recipient.md)** object that is identified by the given entry ID.


## Syntax

_expression_. `GetRecipientFromID`( `_EntryID_` )

_expression_ A variable that represents an '[Account](Outlook.Account.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _EntryID_|Required| **String**|The **[EntryID](Outlook.Recipient.EntryID.md)** of the recipient.|

## Return value

A **Recipient** object that represents the recipient associated with the specified entry ID.


## Remarks

This method is similar to the  **[GetRecipientFromID](Outlook.NameSpace.GetRecipientFromID.md)** method of the **[NameSpace](Outlook.NameSpace.md)** object. If there are multiple Microsoft Exchange accounts in the current profile, use the **GetRecipientFromID** method for the corresponding account.


## See also


[Account Object](Outlook.Account.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]