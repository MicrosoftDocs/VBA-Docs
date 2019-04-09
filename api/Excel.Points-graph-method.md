---
title: Points method (Excel Graph)
keywords: vbagr10.chm3077627
f1_keywords:
- vbagr10.chm3077627
ms.prod: excel
api_name:
- Excel.Points
ms.assetid: d5ec1f3d-a530-d967-4809-850dae59e5be
ms.date: 04/09/2019
localization_priority: Normal
---


# Points method (Excel Graph)

Returns an object that represents a single point or a collection of all the points in the series. Read-only.

## Syntax

_expression_.**Points** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ | Optional |**Variant**| The name or number of the point.|

## Example

This example applies a data label to point one in series one.

```vb
myChart.SeriesCollection(1).Points(1).ApplyDataLabels
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]