---
title: Series.Smooth property (Excel)
keywords: vbaxl10.chm578106
f1_keywords:
- vbaxl10.chm578106
ms.prod: excel
api_name:
- Excel.Series.Smooth
ms.assetid: 24cb3fc6-a6ed-71ca-1aab-c1ea23480d00
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.Smooth property (Excel)

**True** if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts. Read/write.


## Syntax

_expression_.**Smooth**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Example

This example turns on curve smoothing for series one on Chart1. The example should be run on a 2D line chart.

```vb
Charts("Chart1").SeriesCollection(1).Smooth = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]