---
title: TaskRequestItem.Close Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.Close
ms.assetid: d572bebe-11e5-9525-ce99-f4eb33255410
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.Close Event (Outlook)

Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.


## Syntax

_expression_. `Close`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestItem](./Outlook.TaskRequestItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the close operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the close operation isn't completed and the inspector is left open.

If you use the  **[Close](Outlook.Inspector.Close(method).md)** method to fire this event, it can only be canceled if the **Close** method uses the **olPromptForSave** argument.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

