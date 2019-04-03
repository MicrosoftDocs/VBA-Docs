---
title: ContactItem.ReplyAll event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.ReplyAll
ms.assetid: 380f187f-e914-5810-baaf-07473f1719f1
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.ReplyAll event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `ReplyAll`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](Outlook.MailItem.md)** object.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]