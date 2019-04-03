---
title: JournalItem.Send event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.Send
ms.assetid: 0de70191-07cf-fac9-fa0c-49e409e964ac
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.Send event (Outlook)

Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Send`( `_Cancel_` )

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the item is not sent.


## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]