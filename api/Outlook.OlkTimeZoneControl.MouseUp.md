---
title: OlkTimeZoneControl.MouseUp event (Outlook)
keywords: vbaol11.chm1000520
f1_keywords:
- vbaol11.chm1000520
ms.prod: outlook
api_name:
- Outlook.OlkTimeZoneControl.MouseUp
ms.assetid: 93dc1208-11c6-5afc-27d0-ca02a5ddcbe6
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeZoneControl.MouseUp event (Outlook)

Occurs after the user releases a mouse button that has been pressed on the control.


## Syntax

_expression_.**MouseUp** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents an [OlkTimeZoneControl](Outlook.OlkTimeZoneControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **OlMouseButton**|An **[OlMouseButton](Outlook.OlMouseButton.md)** constant that specifies which button on the mouse has been pressed.|
| _Shift_|Required| **OlShiftState**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
| _X_|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
| _Y_|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## See also


[OlkTimeZoneControl Object](Outlook.OlkTimeZoneControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]