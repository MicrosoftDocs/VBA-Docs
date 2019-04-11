---
title: ColumnGroups method (Excel Graph)
keywords: vbagr10.chm65547
f1_keywords:
- vbagr10.chm65547
ms.prod: excel
api_name:
- Excel.ColumnGroups
ms.assetid: dcb4d7e0-ce56-46d9-35d9-d9653bbb6f97
ms.date: 04/09/2019
localization_priority: Normal
---


# ColumnGroups method (Excel Graph)

On a 2D chart, returns an object that represents either a single column chart group or a collection of the column chart groups.

## Syntax

_expression_.**ColumnGroups** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **[ChartGroups](excel.chartgroups(collection).md)** collection.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Optional |**Variant**|The index number of the specified column chart group.|

## Example

This example sets the space between column clusters in the 2D column chart group to be 50 percent of the column width.

```vb
myChart.ColumnGroups(1).GapWidth = 50
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]