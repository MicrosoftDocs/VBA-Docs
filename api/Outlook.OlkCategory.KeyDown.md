---
title: OlkCategory.KeyDown Event (Outlook)
keywords: vbaol11.chm1000456
f1_keywords:
- vbaol11.chm1000456
ms.prod: outlook
api_name:
- Outlook.OlkCategory.KeyDown
ms.assetid: dcaaff84-eb0a-77a7-998d-3327cc7d02bc
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkCategory.KeyDown Event (Outlook)

Occurs when a user presses a key.


## Syntax

_expression_. `KeyDown`( `_KeyCode_` , `_Shift_` )

_expression_ A variable that represents an [OlkCategory](./Outlook.OlkCategory.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyDown** event is accessible through the _Shift_ parameter.


## See also


[OlkCategory Object](Outlook.OlkCategory.md)

