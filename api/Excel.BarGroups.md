---
title: BarGroups method (Excel Graph)
keywords: vbagr10.chm65546
f1_keywords:
- vbagr10.chm65546
ms.prod: excel
api_name:
- Excel.BarGroups
ms.assetid: a00e484e-05ec-2eaa-cc33-05b77a4af0b5
ms.date: 04/06/2019
localization_priority: Normal
---


# BarGroups method (Excel Graph)

On a 2D chart, this method returns an object that represents either a single bar chart group or a collection of all the bar chart groups.

## Syntax

_expression_.**BarGroups** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **[ChartGroups](excel.chartgroups(collection).md)** collection.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Optional |**Variant**|The index number of the specified bar chart group.|

## Example

This example sets the space between bar clusters in the 2D bar chart group to be 50 percent of the bar width.

```vb
myChart.BarGroups(1).GapWidth = 50
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]