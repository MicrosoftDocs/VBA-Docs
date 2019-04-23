---
title: SchemeColor property (Excel Graph)
keywords: vbagr10.chm5207954
f1_keywords:
- vbagr10.chm5207954
ms.prod: excel
api_name:
- Excel.SchemeColor
ms.assetid: a90b4570-dae3-4ca1-563a-0467efbf9bca
ms.date: 04/12/2019
localization_priority: Normal
---


# SchemeColor property (Excel Graph)

Returns or sets the color of the specified **[ChartColorFormat](excel.chartcolorformat.md)** object as an index in the current color scheme. Read/write **Long**.


## Syntax

_expression_.**SchemeColor**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the foreground color, background color, and gradient for the chart area fill on the chart.

```vb
With myChart.ChartArea.Fill 
 .Visible = True 
 .ForeColor.SchemeColor = 15 
 .BackColor.SchemeColor = 17 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]