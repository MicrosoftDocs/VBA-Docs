---
title: MinorTickMark property (Excel Graph)
keywords: vbagr10.chm65563
f1_keywords:
- vbagr10.chm65563
ms.prod: excel
api_name:
- Excel.MinorTickMark
ms.assetid: cbb515d8-fdae-2546-f13b-80ed75cc4192
ms.date: 04/11/2019
localization_priority: Normal
---


# MinorTickMark property (Excel Graph)

Returns or sets the type of minor tick mark for the specified axis. Read/write **[XlTickMark](excel.xltickmark.md)**.

## Syntax

_expression_.**MinorTickMark**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example sets the minor tick marks for the value axis to be inside the axis.

```vb
myChart.Axes(xlValue).MinorTickMark = xlTickMarkInside
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]