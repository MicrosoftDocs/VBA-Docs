---
title: LineGroups method (Excel Graph)
keywords: vbagr10.chm3077623
f1_keywords:
- vbagr10.chm3077623
api_name:
- Excel.LineGroups
ms.assetid: 3a8083b5-8b71-e28b-c775-6be50544d6b2
ms.date: 04/09/2019
ms.localizationpriority: medium
---


# LineGroups method (Excel Graph)

On a 2D chart, returns an object that represents either a single line chart group or a collection of the line chart groups.

## Syntax

_expression_.**LineGroups** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **[ChartGroups](excel.chartgroups(collection).md)** collection.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ | Optional |**Variant**| Specifies the chart group.|

## Example

This example sets line group one to use a different color for each data marker. The example should be run on a 2D chart.

```vb
myChart.LineGroups(1).VaryByCategories = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]