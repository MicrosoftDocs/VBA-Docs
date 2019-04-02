---
title: OlkTimeControl.KeyPress event (Outlook)
keywords: vbaol11.chm1000409
f1_keywords:
- vbaol11.chm1000409
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl.KeyPress
ms.assetid: 58294e95-6774-e32f-22dd-4dea1e28afc6
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeControl.KeyPress event (Outlook)

Occurs when the user presses an ANSI key.


## Syntax

_expression_. `KeyPress`( `_KeyAscii_` )

_expression_ A variable that represents an [OlkTimeControl](Outlook.OlkTimeControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyAscii_|Required| **Long**|The numerical value of the key pressed.|

## Remarks

An ANSI key is one that produces a typeable character when the user presses it. The  **KeyPress** event occurs when the user presses an ANSI key on a running form while the form or a control on it has the focus. The event can occur either before or after the key is released.

A  **KeyPress** event does not occur under the following conditions:


- Pressing  **TAB**
    
- Pressing  **ENTER**
    
- Pressing an arrow key
    
- When a keystroke causes the focus to move from one control to another.
    



## See also


[OlkTimeControl Object](Outlook.OlkTimeControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]