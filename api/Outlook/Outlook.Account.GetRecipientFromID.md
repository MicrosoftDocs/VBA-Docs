---
title: Account.GetRecipientFromID Method (Outlook)
keywords: vbaol11.chm3428
f1_keywords:
- vbaol11.chm3428
ms.prod: outlook
api_name:
- Outlook.Account.GetRecipientFromID
ms.assetid: 7b97ce67-6015-ece6-de1b-6d4226be83aa
ms.date: 06/08/2017
---


# Account.GetRecipientFromID Method (Outlook)

Returns the **[Recipient](Outlook.Recipient.md)** object that is identified by the given entry ID.


## Syntax

 _expression_ . **GetRecipientFromID**( **_EntryID_** )

 _expression_ A variable that represents an **[Account](Outlook.Account.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EntryID_|Required| **String**|The  **[EntryID](Outlook.Recipient.EntryID.md)** of the recipient.|

### Return Value

A  **Recipient** object that represents the recipient associated with the specified entry ID.


## Remarks

This method is similar to the  **[GetRecipientFromID](Outlook.NameSpace.GetRecipientFromID.md)** method of the **[NameSpace](Outlook.NameSpace.md)** object. If there are multiple Microsoft Exchange accounts in the current profile, use the **GetRecipientFromID** method for the corresponding account.


## See also


#### Concepts


[Account Object](Outlook.Account.md)

