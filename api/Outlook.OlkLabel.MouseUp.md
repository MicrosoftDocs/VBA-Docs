---
title: OlkLabel.MouseUp event (Outlook)
keywords: vbaol11.chm1000106
f1_keywords:
- vbaol11.chm1000106
ms.prod: outlook
api_name:
- Outlook.OlkLabel.MouseUp
ms.assetid: edc01865-1bbd-5e87-a857-3e91b28167d2
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkLabel.MouseUp event (Outlook)

Occurs after the user releases a mouse button that has been pressed on the control.


## Syntax

_expression_.**MouseUp** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents an [OlkLabel](Outlook.OlkLabel.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Integer**|An  **[OlMouseButton](Outlook.OlMouseButton.md)** constant that specifies which button on the mouse has been pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
| _X_|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
| _Y_|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## See also


[OlkLabel Object](Outlook.OlkLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]