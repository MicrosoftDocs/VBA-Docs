---
title: OlkTimeZoneControl.MouseMove event (Outlook)
keywords: vbaol11.chm1000519
f1_keywords:
- vbaol11.chm1000519
ms.prod: outlook
api_name:
- Outlook.OlkTimeZoneControl.MouseMove
ms.assetid: 3de1bd35-3351-d70d-9fa4-d90f7d059f87
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeZoneControl.MouseMove event (Outlook)

Occurs after a mouse movement has been registered over the control.


## Syntax

_expression_.**MouseMove** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents an [OlkTimeZoneControl](Outlook.OlkTimeZoneControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **OlMouseButton**|An  **[OlMouseButton](Outlook.OlMouseButton.md)** constant that specifies which button on the mouse has been pressed.|
| _Shift_|Required| **OlShiftState**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
| _X_|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
| _Y_|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## Remarks

Pressing the  **ALT** key fires the **MouseMove** event.


## See also


[OlkTimeZoneControl Object](Outlook.OlkTimeZoneControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]