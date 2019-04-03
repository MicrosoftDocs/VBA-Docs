---
title: OlkContactPhoto.KeyDown event (Outlook)
keywords: vbaol11.chm1000318
f1_keywords:
- vbaol11.chm1000318
ms.prod: outlook
api_name:
- Outlook.OlkContactPhoto.KeyDown
ms.assetid: 5ec4abe0-5600-ea94-c7a8-5f46d4ac587a
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkContactPhoto.KeyDown event (Outlook)

Occurs when a user presses a key.


## Syntax

_expression_. `KeyDown`( `_KeyCode_` , `_Shift_` )

_expression_ A variable that represents an [OlkContactPhoto](Outlook.OlkContactPhoto.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyDown** event is accessible through the _Shift_ parameter.


## See also


[OlkContactPhoto Object](Outlook.OlkContactPhoto.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]