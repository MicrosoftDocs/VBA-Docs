---
title: Orientation property (Excel Graph)
keywords: vbagr10.chm65670
f1_keywords:
- vbagr10.chm65670
ms.prod: excel
ms.assetid: 1e4e111c-5144-a509-4791-e8ca31c3de5e
ms.date: 04/11/2019
localization_priority: Normal
---


# Orientation property (Excel Graph)

Returns or sets the text orientation. Can be an integer value from -90 degrees to 90 degrees, or one of the **[XlOrientation](excel.xlorientation.md)** constants. Read/write **[XlTickLabelOrientation](excel.xlticklabelorientation.md)** for all objects, except for the **[TickLabels](excel.ticklabels-graph-object.md)** object, which is read/write **Variant**.

## Syntax

_expression_.**Orientation**

_expression_ Required. An expression that returns one of the above objects.

## Example

This example sets the orientation for the chart title.

```vb
myChart.ChartTitle.Orientation = xlHorizontal
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]