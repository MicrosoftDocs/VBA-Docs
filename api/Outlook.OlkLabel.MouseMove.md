---
title: OlkLabel.MouseMove event (Outlook)
keywords: vbaol11.chm1000105
f1_keywords:
- vbaol11.chm1000105
ms.prod: outlook
api_name:
- Outlook.OlkLabel.MouseMove
ms.assetid: b956ea54-96cb-d5d1-e542-ea9f464117c7
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkLabel.MouseMove event (Outlook)

Occurs after a mouse movement has been registered over the control.


## Syntax

_expression_.**MouseMove** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents an [OlkLabel](Outlook.OlkLabel.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Integer**|An **[OlMouseButton](Outlook.OlMouseButton.md)** constant that specifies which button on the mouse has been pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](Outlook.OlShiftState.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
| _X_|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
| _Y_|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## Remarks

Pressing the  **ALT** key fires the **MouseMove** event.


## See also


[OlkLabel Object](Outlook.OlkLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]