---
title: OlkCategory.KeyPress event (Outlook)
keywords: vbaol11.chm1000457
f1_keywords:
- vbaol11.chm1000457
ms.prod: outlook
api_name:
- Outlook.OlkCategory.KeyPress
ms.assetid: 15a26c77-d82f-f734-f294-54d38a657af3
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkCategory.KeyPress event (Outlook)

Occurs when the user presses an ANSI key.


## Syntax

_expression_. `KeyPress`( `_KeyAscii_` )

_expression_ A variable that represents an [OlkCategory](Outlook.OlkCategory.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyAscii_|Required| **Long**|The numerical value of the key pressed.|

## Remarks

An ANSI key is one that produces a typeable character when the user presses it. The **KeyPress** event occurs when the user presses an ANSI key on a running form while the form or a control on it has the focus. The event can occur either before or after the key is released.

A **KeyPress** event does not occur under the following conditions:


- Pressing  **TAB**
    
- Pressing  **ENTER**
    
- Pressing an arrow key
    
- When a keystroke causes the focus to move from one control to another.
    



## See also


[OlkCategory Object](Outlook.OlkCategory.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]