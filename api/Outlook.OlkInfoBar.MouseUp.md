---
title: OlkInfoBar.MouseUp event (Outlook)
keywords: vbaol11.chm1000303
f1_keywords:
- vbaol11.chm1000303
ms.prod: outlook
api_name:
- Outlook.OlkInfoBar.MouseUp
ms.assetid: daff2dbd-0da7-e5b0-7425-8aaf325b4b8a
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkInfoBar.MouseUp event (Outlook)

Occurs after the user releases a mouse button that has been pressed on the control.


## Syntax

_expression_.**MouseUp** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents an [OlkInfoBar](Outlook.OlkInfoBar.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Integer**|An **[OlMouseButton](Outlook.OlMouseButton.md)** constant that specifies which button on the mouse has been pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
| _X_|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
| _Y_|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## See also


[OlkInfoBar Object](Outlook.OlkInfoBar.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]