---
title: ChartGroups method (Excel Graph)
keywords: vbagr10.chm65544
f1_keywords:
- vbagr10.chm65544
ms.prod: excel
api_name:
- Excel.ChartGroups
ms.assetid: e25258c1-14d4-bb0c-b442-f6c811b19847
ms.date: 04/06/2019
localization_priority: Normal
---


# ChartGroups method (Excel Graph)

Returns an object that represents either a single chart group or a collection of all the chart groups in the chart. The returned collection includes every type of group.

## Syntax

_expression_.**ChartGroups** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Optional |**Variant**|The chart group number.|

## Example

This example turns on up and down bars for chart group one and then sets their colors. The example should be run on a 2D line chart containing two series that intersect at one or more data points.

```vb
With myChart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]