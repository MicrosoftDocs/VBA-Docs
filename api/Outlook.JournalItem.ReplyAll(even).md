---
title: JournalItem.ReplyAll event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.ReplyAll
ms.assetid: 86ab09f8-92f5-320e-9ec0-3be1f63c4583
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.ReplyAll event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item, or when the **ReplyAll** method is called for the item, which is an instance of the parent object.


## Syntax

_expression_. `ReplyAll`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a **[MailItem](Outlook.MailItem.md)** object.


## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]