---
title: HasAxis property (Excel Graph)
keywords: vbagr10.chm65588
f1_keywords:
- vbagr10.chm65588
ms.prod: excel
api_name:
- Excel.HasAxis
ms.assetid: 2de3c3a1-7b9c-a4d9-40cb-906fd5d6f4cb
ms.date: 04/11/2019
localization_priority: Normal
---


# HasAxis property (Excel Graph)

Returns or sets which axes exist on the chart. Read/write **Variant**.

## Syntax

_expression_.**HasAxis** (_Index1_, _Index2_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index1_ |Optional |**[XlAxisType](excel.xlaxistype.md)**|The type of axis. Can be one of these **XlAxisType** constants: **xlCategory**, **xlValue**, or **xlSeriesAxis**. Series axes apply only to 3D charts.|
|_Index2_ |Optional |**[XlAxisGroup](excel.xlaxisgroup.md)**|The axis priority. Can be one of these **XlAxisGroup** constants: **xlPrimary** or **xlSecondary**. 3D charts have only one set of axes.|

## Remarks

Graph may create or delete axes if you change the chart type or change the **[AxisGroup](Excel.AxisGroup.md)** property.


## Example

This example turns on the primary value axis.

```vb
myChart.HasAxis(xlValue, xlPrimary) = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]