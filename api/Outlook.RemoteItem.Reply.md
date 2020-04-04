---
title: RemoteItem.Reply event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.Reply
ms.assetid: 47b49c1a-2e70-0265-d36d-58cf3800ffaf
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.Reply event (Outlook)

Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Reply`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the reply operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a **[MailItem](Outlook.MailItem.md)** object.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the reply action is not completed and the new item is not displayed.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]