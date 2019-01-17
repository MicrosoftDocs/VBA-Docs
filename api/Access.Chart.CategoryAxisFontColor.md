---
title: Chart.CategoryAxisFontColor property (Access)
keywords: vbaac10.chm6134
f1_keywords:
- vbaac10.chm6134
ms.prod: access
api_name:
- Access.Chart.CategoryAxisFontColor
ms.date: 11/28/2018
localization_priority: Normal
---


# Chart.CategoryAxisFontColor property (Access)

Returns or sets the font color used by the category axis. Read/write **Long**.

You can use a **[system color constant](../language/reference/user-interface-help/system-color-constants.md)** or the RGB function to set a color programmatically as shown in the following example. You can also browse and select a color from the Design View palette.

## Syntax

_expression_.**CategoryAxisFontColor**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.


## Example

In this example, the **CategoryAxisFontColor** is initially set to a system color constant before it is changed to an RGB value.

```vb
With myChart
 MsgBox ("Applying a system color constant")
 .CategoryAxisFontColor = vbHighlight
 MsgBox ("Applying an RGB value")
 .CategoryAxisFontColor = RGB(255, 165, 0)
End With
```

