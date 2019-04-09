---
title: Solid method (Excel Graph)
keywords: vbagr10.chm3077634
f1_keywords:
- vbagr10.chm3077634
ms.prod: excel
api_name:
- Excel.Solid
ms.assetid: 34fcc8d7-df60-2bad-0674-a1b9819509f7
ms.date: 04/09/2019
localization_priority: Normal
---


# Solid method (Excel Graph)

Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.

## Syntax

_expression_.**Solid**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example converts the chart area fill to a solid color.

```vb
myChart.ChartArea.Fill.Solid
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]