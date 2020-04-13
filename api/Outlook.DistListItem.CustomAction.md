---
title: DistListItem.CustomAction event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.CustomAction
ms.assetid: cdd8a5cb-1a18-85fe-0a26-b89c6bb61a7d
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.CustomAction event (Outlook)

Occurs when a custom action of an item (which is an instance of the parent object) executes.


## Syntax

_expression_. `CustomAction`( `_Action_` , `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Action_|Required| **Object**|The **[Action](Outlook.Action.md)** object.|
| _Response_|Required| **Object**|The newly created item resulting from the custom action.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the custom action is not completed.|

## Remarks

The **Action** object and the newly created item resulting from the custom action are passed to the event.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the custom action operation is not completed.


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]