---
title: SeriesCollection method (Excel Graph)
keywords: vbagr10.chm65604
f1_keywords:
- vbagr10.chm65604
ms.prod: excel
api_name:
- Excel.SeriesCollection
ms.assetid: 9de760ab-0a6d-7cab-e378-b7f341f5b87d
ms.date: 04/09/2019
localization_priority: Normal
---


# SeriesCollection method (Excel Graph)

Returns an object that represents either a single series or a collection of all the series in the chart or chart group.

## Syntax

_expression_.**SeriesCollection** (_Index_)

_expression_ Required. An expression that returns the **[SeriesCollection](excel.seriescollection(collection).md)** collection. 

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ | Optional |**Variant**| The name or number of the series.|

## Example

This example turns on data labels for series one.

```vb
myChart.SeriesCollection(1).HasDataLabels = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]