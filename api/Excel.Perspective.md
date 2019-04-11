---
title: Perspective property (Excel Graph)
keywords: vbagr10.chm65593
f1_keywords:
- vbagr10.chm65593
ms.prod: excel
api_name:
- Excel.Perspective
ms.assetid: 84ddaf6c-1204-1a7b-55e5-7d3cf2787a2c
ms.date: 04/11/2019
localization_priority: Normal
---


# Perspective property (Excel Graph)

Returns or sets the perspective for the 3D chart view. Must be from 0 through 100. This property is ignored if the **[RightAngleAxes](Excel.RightAngleAxes.md)** property is **True**. Read/write **Long**.

## Syntax

_expression_.**Perspective**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the perspective of _myChart_ to 70. The example should be run on a 3D chart.

```vb
myChart.RightAngleAxes = False 
myChart.Perspective = 70
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]