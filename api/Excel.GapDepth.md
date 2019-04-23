---
title: GapDepth property (Excel Graph)
keywords: vbagr10.chm5207398
f1_keywords:
- vbagr10.chm5207398
ms.prod: excel
api_name:
- Excel.GapDepth
ms.assetid: 0aa59fe6-29bf-c014-8c11-18481f9c5603
ms.date: 04/10/2019
localization_priority: Normal
---


# GapDepth property (Excel Graph)

Returns or sets the distance between the data series on a 3D chart, as a percentage of the marker width. The value of this property must be between 0 and 500. Read/write **Long**.

## Syntax

_expression_.**GapDepth**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the distance between the data series to 200 percent of the marker width. The example should be run on a 3D chart (the **GapDepth** property fails on 2D charts).

```vb
myChart.GapDepth = 200
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]