---
title: OlkDateControl.KeyUp Event (Outlook)
keywords: vbaol11.chm1000371
f1_keywords:
- vbaol11.chm1000371
ms.prod: outlook
api_name:
- Outlook.OlkDateControl.KeyUp
ms.assetid: 7776832b-fdb0-cd2b-efa3-97dab74065e6
ms.date: 06/08/2017
---


# OlkDateControl.KeyUp Event (Outlook)

Occurs when the user releases a key.


## Syntax

 _expression_. `KeyUp`( `_KeyCode_` , `_Shift_` )

 _expression_ A variable that represents an [OlkDateControl](./Outlook.OlkDateControl.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys (**SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyUp** event is accessible through the _Shift_ parameter.


## See also


[OlkDateControl Object](Outlook.OlkDateControl.md)

