---
title: OlkOptionButton.KeyDown event (Outlook)
keywords: vbaol11.chm1000186
f1_keywords:
- vbaol11.chm1000186
ms.prod: outlook
api_name:
- Outlook.OlkOptionButton.KeyDown
ms.assetid: f236a9a0-cbde-d6f6-8fe8-681543de9aa5
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkOptionButton.KeyDown event (Outlook)

Occurs when a user presses a key.


## Syntax

_expression_. `KeyDown`( `_KeyCode_` , `_Shift_` )

_expression_ A variable that represents an [OlkOptionButton](Outlook.OlkOptionButton.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyDown** event is accessible through the _Shift_ parameter.


## See also


[OlkOptionButton Object](Outlook.OlkOptionButton.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]