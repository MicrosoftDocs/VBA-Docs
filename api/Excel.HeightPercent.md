---
title: HeightPercent property (Excel Graph)
keywords: vbagr10.chm5207525
f1_keywords:
- vbagr10.chm5207525
ms.prod: excel
api_name:
- Excel.HeightPercent
ms.assetid: 711c65bd-5603-2678-e07b-fa20d55ada4b
ms.date: 04/11/2019
localization_priority: Normal
---


# HeightPercent property (Excel Graph)

Returns or sets the height of a 3D chart as a percentage of the chart width (between 5 and 500 percent). Read/write **Long**.

## Syntax

_expression_.**HeightPercent**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the height of the chart to 80 percent of its width. The example should be run on a 3D chart.

```vb
myChart.HeightPercent = 80
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]