---
title: OlkComboBox.KeyUp event (Outlook)
keywords: vbaol11.chm1000244
f1_keywords:
- vbaol11.chm1000244
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.KeyUp
ms.assetid: 22f2f29c-f4ea-764a-85a0-90d11becf5dc
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox.KeyUp event (Outlook)

Occurs when the user releases a key.


## Syntax

_expression_. `KeyUp`( `_KeyCode_` , `_Shift_` )

_expression_ A variable that represents an [OlkComboBox](Outlook.OlkComboBox.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyUp** event is accessible through the _Shift_ parameter.


## See also


[OlkComboBox Object](Outlook.OlkComboBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]