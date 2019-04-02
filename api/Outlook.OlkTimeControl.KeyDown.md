---
title: OlkTimeControl.KeyDown event (Outlook)
keywords: vbaol11.chm1000408
f1_keywords:
- vbaol11.chm1000408
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl.KeyDown
ms.assetid: 1214ffd2-033e-13bb-309e-254d98f903c0
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeControl.KeyDown event (Outlook)

Occurs when a user presses a key.


## Syntax

_expression_. `KeyDown`( `_KeyCode_` , `_Shift_` )

_expression_ A variable that represents an [OlkTimeControl](Outlook.OlkTimeControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyDown** event is accessible through the _Shift_ parameter.


## See also


[OlkTimeControl Object](Outlook.OlkTimeControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]