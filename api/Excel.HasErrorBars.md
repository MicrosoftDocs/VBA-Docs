---
title: HasErrorBars property (Excel Graph)
keywords: vbagr10.chm65696
f1_keywords:
- vbagr10.chm65696
ms.prod: excel
api_name:
- Excel.HasErrorBars
ms.assetid: f16a2ffe-b481-ec32-1144-8c1e5718243f
ms.date: 04/11/2019
localization_priority: Normal
---


# HasErrorBars property (Excel Graph)

**True** if the series has error bars. This property isn't available for 3D charts. Read/write **Boolean**.

## Syntax

_expression_.**HasErrorBars**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example removes error bars from series one. The example should be run on a 2D line chart that has error bars for series one.

```vb
myChart.SeriesCollection(1).HasErrorBars = False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]