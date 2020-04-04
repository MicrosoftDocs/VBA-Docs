---
title: TaskRequestAcceptItem.Reply event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.Reply
ms.assetid: f06a8e69-b674-7b70-e7a4-7f234ebb5b1d
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.Reply event (Outlook)

Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Reply`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the reply operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a **[MailItem](Outlook.MailItem.md)** object.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the reply action is not completed and the new item is not displayed.


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]