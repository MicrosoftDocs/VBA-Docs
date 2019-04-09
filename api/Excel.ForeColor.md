---
title: ForeColor property (Excel Graph)
keywords: vbagr10.chm5207390
f1_keywords:
- vbagr10.chm5207390
ms.prod: excel
api_name:
- Excel.ForeColor
ms.assetid: 1c1eb700-672e-095d-826c-28cdb7e9de40
ms.date: 04/10/2019
localization_priority: Normal
---


# ForeColor property (Excel Graph)

Returns a **ChartColorFormat** object that represents the foreground fill color.

## Syntax

_expression_.**ForeColor**

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