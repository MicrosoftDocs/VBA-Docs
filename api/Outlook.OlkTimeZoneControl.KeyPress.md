---
title: OlkTimeZoneControl.KeyPress event (Outlook)
keywords: vbaol11.chm1000524
f1_keywords:
- vbaol11.chm1000524
ms.prod: outlook
api_name:
- Outlook.OlkTimeZoneControl.KeyPress
ms.assetid: 4b6f04be-85c2-70f8-001f-30f008fb9b4a
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeZoneControl.KeyPress event (Outlook)

Occurs when the user presses an ANSI key.


## Syntax

_expression_. `KeyPress`( `_KeyAscii_` )

_expression_ A variable that represents an [OlkTimeZoneControl](Outlook.OlkTimeZoneControl.md) object.


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
    
- When a keystroke causes the focus to move from one control to another
    



## See also


[OlkTimeZoneControl Object](Outlook.OlkTimeZoneControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]