---
title: BackColor property (Excel Graph)
keywords: vbagr10.chm67174
f1_keywords:
- vbagr10.chm67174
ms.prod: excel
ms.assetid: 29f8617f-71a2-fa0b-89c7-8b20ff8cd87d
ms.date: 04/09/2019
localization_priority: Normal
---


# BackColor property (Excel Graph)

Returns a **ChartColorFormat** object that represents the fill background color.

## Syntax

_expression_.**BackColor**

_expression_ Required. An expression that returns a **[ChartColorFormat](excel.chartcolorformat.md)** object.

## Example

This example sets the gradient, background color, and foreground color for the chart area fill.

```vb
With myChart.ChartArea.Fill 
 .Visible = True 
 .ForeColor.SchemeColor = 15 
 .BackColor.SchemeColor = 17 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]