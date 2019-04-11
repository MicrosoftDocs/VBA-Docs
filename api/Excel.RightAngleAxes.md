---
title: RightAngleAxes property (Excel Graph)
keywords: vbagr10.chm3077581
f1_keywords:
- vbagr10.chm3077581
ms.prod: excel
api_name:
- Excel.RightAngleAxes
ms.assetid: 5c34e5b4-a936-70a5-cd0c-d9a7a091e8d0
ms.date: 04/12/2019
localization_priority: Normal
---


# RightAngleAxes property (Excel Graph)

**True** if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3D line, column, and bar charts. Read/write **Variant**.

## Syntax

_expression_.**RightAngleAxes**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

If this property is **True**, the **[Perspective](Excel.Perspective.md)** property is ignored.


## Example

This example sets the axes to intersect at right angles. The example should be run on a 3D chart.

```vb
myChart.RightAngleAxes = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]