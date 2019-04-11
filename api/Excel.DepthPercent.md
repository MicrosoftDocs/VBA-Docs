---
title: DepthPercent property (Excel Graph)
keywords: vbagr10.chm5207305
f1_keywords:
- vbagr10.chm5207305
ms.prod: excel
api_name:
- Excel.DepthPercent
ms.assetid: b8c8f784-bc30-cc20-604d-5627b570c1f2
ms.date: 04/10/2019
localization_priority: Normal
---


# DepthPercent property (Excel Graph)

Returns or sets the depth of a 3D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write **Long**.

## Syntax

_expression_.**DepthPercent**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the depth of the chart to be 50 percent of its width. The example should be run on a 3D chart (the **DepthPercent** property fails on 2D charts).

```vb
myChart.DepthPercent = 50
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]