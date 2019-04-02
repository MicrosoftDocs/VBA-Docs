---
title: OlkTimeControl.KeyUp event (Outlook)
keywords: vbaol11.chm1000410
f1_keywords:
- vbaol11.chm1000410
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl.KeyUp
ms.assetid: b2ff348b-6c94-09b3-e8ee-8eb25ac15ba0
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeControl.KeyUp event (Outlook)

Occurs when the user releases a key.


## Syntax

_expression_. `KeyUp`( `_KeyCode_` , `_Shift_` )

_expression_ A variable that represents an [OlkTimeControl](Outlook.OlkTimeControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyUp** event is accessible through the _Shift_ parameter.


## See also


[OlkTimeControl Object](Outlook.OlkTimeControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]