---
title: OlkTextBox.KeyUp event (Outlook)
keywords: vbaol11.chm1000080
f1_keywords:
- vbaol11.chm1000080
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.KeyUp
ms.assetid: 713fb517-9a64-6919-ec8c-e24b9dc066d1
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTextBox.KeyUp event (Outlook)

Occurs when the user releases a key.


## Syntax

_expression_. `KeyUp`( `_KeyCode_` , `_Shift_` )

_expression_ A variable that represents an [OlkTextBox](Outlook.OlkTextBox.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyUp** event is accessible through the _Shift_ parameter.


## See also


[OlkTextBox Object](Outlook.OlkTextBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]