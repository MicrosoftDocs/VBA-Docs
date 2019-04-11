---
title: AreaGroups method (Excel Graph)
keywords: vbagr10.chm3077607
f1_keywords:
- vbagr10.chm3077607
ms.prod: excel
api_name:
- Excel.AreaGroups
ms.assetid: ec2a4a28-2f10-4f4f-bd91-642bf1b8ebe2
ms.date: 04/06/2019
localization_priority: Normal
---


# AreaGroups method (Excel Graph)

On a 2D chart, this method returns an object that represents a single area chart group or a collection of all the area chart groups.

## Syntax

_expression_.**AreaGroups** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **[ChartGroups](excel.chartgroups(collection).md)** collection.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Optional |**Variant**|The index number of the specified chart group.|

## Example

This example turns on drop lines for the 2D area chart group.

```vb
myChart.AreaGroups(1).HasDropLines = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]