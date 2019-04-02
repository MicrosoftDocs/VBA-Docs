---
title: OlkTimeZoneControl.KeyDown event (Outlook)
keywords: vbaol11.chm1000523
f1_keywords:
- vbaol11.chm1000523
ms.prod: outlook
api_name:
- Outlook.OlkTimeZoneControl.KeyDown
ms.assetid: 155fa627-af2a-9c0a-3f56-fb7318843c16
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeZoneControl.KeyDown event (Outlook)

Occurs when a user presses a key.


## Syntax

_expression_. `KeyDown`( `_KeyCode_` , `_Shift_` )

_expression_ A variable that represents an [OlkTimeZoneControl](Outlook.OlkTimeZoneControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **OlShiftState**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyDown** event is accessible through the _Shift_ parameter.


## See also


[OlkTimeZoneControl Object](Outlook.OlkTimeZoneControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]