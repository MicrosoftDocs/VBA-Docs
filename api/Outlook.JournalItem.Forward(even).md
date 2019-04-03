---
title: JournalItem.Forward event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.Forward
ms.assetid: 0e52ded6-1963-8259-862d-cac610102b10
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.Forward event (Outlook)

Occurs when the user selects the  **Forward** action for an item, or when the **Forward** method is called for the item, which is an instance of the parent object.


## Syntax

_expression_. `Forward`( `_Forward_` , `_Cancel_` )

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Forward_|Required| **Object**|The new item being forwarded.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False**, the forward action is not completed and the new item is not displayed.


## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]