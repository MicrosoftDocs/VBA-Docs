---
title: OlkCommandButton.KeyDown Event (Outlook)
keywords: vbaol11.chm1000127
f1_keywords:
- vbaol11.chm1000127
ms.prod: outlook
api_name:
- Outlook.OlkCommandButton.KeyDown
ms.assetid: 626f3437-4101-06e9-5041-39fedd38b687
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkCommandButton.KeyDown Event (Outlook)

Occurs when a user presses a key.


## Syntax

_expression_. `KeyDown`( `_KeyCode_` , `_Shift_` )

_expression_ A variable that represents an [OlkCommandButton](./Outlook.OlkCommandButton.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyDown** event is accessible through the _Shift_ parameter.


## See also


[OlkCommandButton Object](Outlook.OlkCommandButton.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]