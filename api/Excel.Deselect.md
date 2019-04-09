---
title: Deselect method (Excel Graph)
keywords: vbagr10.chm66656
f1_keywords:
- vbagr10.chm66656
ms.prod: excel
api_name:
- Excel.Deselect
ms.assetid: 928e8efa-4b6a-a1ea-2520-615354c8538a
ms.date: 04/09/2019
localization_priority: Normal
---


# Deselect method (Excel Graph)

Cancels the selection for the chart.

## Syntax

_expression_.**Deselect**

_expression_ Required. An expression that returns a **[Chart](Excel.Chart-graph-object.md)** object.


## Example

This example is equivalent to pressing Esc while working on the chart. The example should be run on a chart that has a component (such as an axis) selected.

```vb
myChart.Deselect
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]