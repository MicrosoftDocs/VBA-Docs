---
title: ContactItem.Reply event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.Reply
ms.assetid: 0560988f-95a1-23f5-67af-f94321d9ff39
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.Reply event (Outlook)

Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Reply`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the reply operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](Outlook.MailItem.md)** object.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the reply action is not completed and the new item is not displayed.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]