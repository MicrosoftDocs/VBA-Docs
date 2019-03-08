---
title: Chart.SecondaryValuesAxisFontColor property (Access)
keywords: vbaac10.chm6133
f1_keywords:
- vbaac10.chm6133
ms.prod: access
api_name:
- Access.Chart.SecondaryValuesAxisFontColor
ms.date: 11/28/2018
localization_priority: Normal
---


# Chart.SecondaryValuesAxisFontColor property (Access)

Returns or sets the font color used by the secondary values axis. Read/write **Long**.

You can use a **[system color constant](../language/reference/user-interface-help/system-color-constants.md)** or the RGB function to set a color programmatically as shown in the example. You can also browse and select a color from the Design view palette.


## Syntax

_expression_.**SecondaryValuesAxisFontColor**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.

## Example

In this example, the **SecondaryValuesAxisFontColor** is initially set to a system color constant before it is changed to an RGB value.

```vb
With myChart
 MsgBox ("Applying a system color constant")
 .SecondaryValuesAxisFontColor = vbHighlight
 MsgBox ("Applying an RGB value")
 .SecondaryValuesAxisFontColor = RGB(255, 165, 0)
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]