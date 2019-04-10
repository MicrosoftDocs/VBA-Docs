---
title: Pattern property (Excel Graph)
keywords: vbagr10.chm65631
f1_keywords:
- vbagr10.chm65631
ms.prod: excel
api_name:
- Excel.Pattern
ms.assetid: 3cc8475d-dc65-b2eb-e1ba-2bd95c5c0b03
ms.date: 04/11/2019
localization_priority: Normal
---


# Pattern property (Excel Graph)

For the **[ChartFillFormat](excel.chartfillformat.md)** object, returns or sets the fill pattern. Read-only **[MsoPatternType](office.msopatterntype.md)**. 

For the **[Interior](excel.interior-graph-object.md)** object, returns or sets the interior pattern. Read/write **Variant**.

## Syntax

_expression_.**Pattern**

_expression_ Required. An expression that returns one of the above objects.

## Example

This example adds a crisscross pattern to the interior of the plot area.

```vb
myChart.PlotArea.Interior.Pattern = xlPatternCrissCross
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]