---
title: ChartGroup.Overlap property (Excel)
keywords: vbaxl10.chm568086
f1_keywords:
- vbaxl10.chm568086
ms.prod: excel
api_name:
- Excel.ChartGroup.Overlap
ms.assetid: 6ea1de1a-ecb4-d920-fc34-ed3bf3a767b4
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.Overlap property (Excel)

Specifies how bars and columns are positioned. Can be a value between - 100 and 100. Applies only to 2-D bar and 2-D column charts. Read/write  **Long**.


## Syntax

_expression_. `Overlap`

_expression_ A variable that represents a [ChartGroup](Excel.ChartGroup-graph-object.md) object.


## Remarks

If this property is set to - 100, bars are positioned so that there's one bar width between them. If the overlap is 0 (zero), there's no space between bars (one bar starts immediately after the preceding bar). If the overlap is 100, bars are positioned on top of each other.


## Example

This example sets the overlap for chart group one to - 50. The example should be run on a 2-D column chart that has two or more series.


```vb
Charts("Chart1").ChartGroups(1).Overlap = -50
```


## See also


[ChartGroup Object](Excel.ChartGroup(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]