---
title: DoughnutGroups method (Excel Graph)
keywords: vbagr10.chm3077618
f1_keywords:
- vbagr10.chm3077618
ms.prod: excel
api_name:
- Excel.DoughnutGroups
ms.assetid: 41ca4213-c17b-7bba-c357-7ba65fd55d39
ms.date: 04/09/2019
localization_priority: Normal
---


# DoughnutGroups method (Excel Graph)

On a 2D chart, returns an object that represents either a single doughnut chart group or a collection of the doughnut chart groups.

## Syntax

_expression_.**DoughnutGroups** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **[ChartGroups](excel.chartgroups(collection).md)** collection.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Optional |**Variant**|Specifies the chart group.|

## Example

This example sets the starting angle for doughnut group one.

```vb
myChart.DoughnutGroups(1).FirstSliceAngle = 45
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]