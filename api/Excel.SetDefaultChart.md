---
title: SetDefaultChart method (Excel Graph)
keywords: vbagr10.chm65755
f1_keywords:
- vbagr10.chm65755
ms.prod: excel
api_name:
- Excel.SetDefaultChart
ms.assetid: 1afc1023-654b-67cd-aace-bc4b87747520
ms.date: 04/09/2019
localization_priority: Normal
---


# SetDefaultChart method (Excel Graph)

Specifies the name of the chart template that Graph will use when creating new charts.

## Syntax

_expression_.**SetDefaultChart** (_FormatName_, _Gallery_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FormatName_ |Optional |**Variant**| The name of the specified custom autoformat. This name can be a string that denotes the custom autoformat, or it can be the special constant **xlBuiltIn** to specify the built-in chart template.|
|_Gallery_ |Optional |**Variant**||

## Example

This example sets the default chart template to the custom autoformat named Monthly Sales.

```vb
myChart.Application.SetDefaultChart FormatName:="Monthly Sales"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]