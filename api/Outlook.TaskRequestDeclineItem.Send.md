---
title: TaskRequestDeclineItem.Send Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.Send
ms.assetid: e78cf949-6fdf-db40-8638-e23dcb16529c
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.Send Event (Outlook)

Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Send`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestDeclineItem](./Outlook.TaskRequestDeclineItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the item is not sent.


## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

