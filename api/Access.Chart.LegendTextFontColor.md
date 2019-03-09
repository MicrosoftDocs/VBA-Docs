---
title: Chart.LegendTextFontColor property (Access)
keywords: vbaac10.chm6152
f1_keywords:
- vbaac10.chm6152
ms.prod: access
api_name:
- Access.Chart.LegendTextFontColor
ms.date: 11/28/2018
localization_priority: Normal
---


# Chart.LegendTextFontColor property (Access)

Returns or sets the font color used by the chart legend. Read/write **Long**.

You can use a **[system color constant](../language/reference/user-interface-help/system-color-constants.md)** or the RGB function to set a color programmatically as shown in the example. You can also browse and select a color from the Design view palette.


## Syntax

_expression_.**LegendTextFontColor**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.


## Example

In this example, the **LegendTextFontColor** is initially set to a system color constant before it is changed to an RGB value.

```vb
With myChart
 MsgBox ("Applying a system color constant")
 .ChartTitleFontColor = vbHighlight
 MsgBox ("Applying an RGB value")
 .ChartTitleFontColor = RGB(255, 165, 0)
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]