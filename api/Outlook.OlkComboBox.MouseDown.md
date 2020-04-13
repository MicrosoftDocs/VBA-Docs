---
title: OlkComboBox.MouseDown event (Outlook)
keywords: vbaol11.chm1000237
f1_keywords:
- vbaol11.chm1000237
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.MouseDown
ms.assetid: 7d9385c4-8c21-5ff3-ec00-789223f76072
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox.MouseDown event (Outlook)

Occurs when the user presses a mouse button on the control.


## Syntax

_expression_.**MouseDown** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents an [OlkComboBox](Outlook.OlkComboBox.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Integer**|An **[OlMouseButton](Outlook.OlMouseButton.md)** constant that specifies which button on the mouse has been pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
| _X_|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
| _Y_|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## See also


[OlkComboBox Object](Outlook.OlkComboBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]