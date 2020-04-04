---
title: MeetingItem.ReplyAll event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.ReplyAll
ms.assetid: d93bd51d-a169-0007-4188-4fff829dbb1e
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.ReplyAll event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item, or when the **ReplyAll** method is called for the item, which is an instance of the parent object.


## Syntax

_expression_. `ReplyAll`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a **[MailItem](Outlook.MailItem.md)** object.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]