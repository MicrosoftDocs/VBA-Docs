---
title: RGB property (Excel Graph)
keywords: vbagr10.chm5207930
f1_keywords:
- vbagr10.chm5207930
ms.prod: excel
api_name:
- Excel.RGB
ms.assetid: bb3dbad0-a96a-969d-1234-ee9cf59e4c87
ms.date: 04/12/2019
localization_priority: Normal
---


# RGB property (Excel Graph)

Returns the red-green-blue value of the specified color. Read-only **Long**.

## Syntax

_expression_.**RGB**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the color of the legend font to the foreground fill color of the plot area.

```vb
myChart.Legend.Font.Color = _ 
 myChart.PlotArea.Fill.ForeColor.RGB
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]