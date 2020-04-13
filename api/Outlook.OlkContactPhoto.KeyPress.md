---
title: OlkContactPhoto.KeyPress event (Outlook)
keywords: vbaol11.chm1000319
f1_keywords:
- vbaol11.chm1000319
ms.prod: outlook
api_name:
- Outlook.OlkContactPhoto.KeyPress
ms.assetid: 43b7f7e0-79c5-e02c-5d9e-a204098509c2
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkContactPhoto.KeyPress event (Outlook)

Occurs when the user presses an ANSI key.


## Syntax

_expression_. `KeyPress`( `_KeyAscii_` )

_expression_ A variable that represents an [OlkContactPhoto](Outlook.OlkContactPhoto.md) object.


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


[OlkContactPhoto Object](Outlook.OlkContactPhoto.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]