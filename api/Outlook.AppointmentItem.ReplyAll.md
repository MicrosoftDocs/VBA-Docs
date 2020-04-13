---
title: AppointmentItem.ReplyAll event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.ReplyAll
ms.assetid: c49245b9-0770-f551-c382-3f5745aead04
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.ReplyAll event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `ReplyAll`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a **[MailItem](Outlook.MailItem.md)** object.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]