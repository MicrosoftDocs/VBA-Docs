---
title: OlkContactPhoto.MouseMove event (Outlook)
keywords: vbaol11.chm1000314
f1_keywords:
- vbaol11.chm1000314
ms.prod: outlook
api_name:
- Outlook.OlkContactPhoto.MouseMove
ms.assetid: d2f0b94b-4825-c3be-d2b6-070e0fb2ff44
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkContactPhoto.MouseMove event (Outlook)

Occurs after a mouse movement has been registered over the control.


## Syntax

_expression_.**MouseMove** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents an [OlkContactPhoto](Outlook.OlkContactPhoto.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Integer**|An  **[OlMouseButton](Outlook.OlMouseButton.md)** constant that specifies which button on the mouse has been pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
| _X_|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
| _Y_|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## Remarks

Pressing the  **ALT** key fires the **MouseMove** event.


## See also


[OlkContactPhoto Object](Outlook.OlkContactPhoto.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]