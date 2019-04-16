---
title: OlkListBox.MouseDown event (Outlook)
keywords: vbaol11.chm1000282
f1_keywords:
- vbaol11.chm1000282
ms.prod: outlook
api_name:
- Outlook.OlkListBox.MouseDown
ms.assetid: d2ff81b0-6875-0b2a-46c1-4fd6ff2bb42c
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkListBox.MouseDown event (Outlook)

Occurs when the user presses a mouse button on the control.


## Syntax

_expression_.**MouseDown** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents an [OlkListBox](Outlook.OlkListBox.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Integer**|An  **[OlMouseButton](Outlook.OlMouseButton.md)** constant that specifies which button on the mouse has been pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
| _X_|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
| _Y_|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## See also


[OlkListBox Object](Outlook.OlkListBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]