---
title: Axes method (Excel Graph)
keywords: vbagr10.chm3077608
f1_keywords:
- vbagr10.chm3077608
ms.prod: excel
api_name:
- Excel.Axes
ms.assetid: 040bf3e2-f60f-935b-9803-6f9bf146bee7
ms.date: 04/06/2019
localization_priority: Normal
---


# Axes method (Excel Graph)

Returns an object that represents either a single axis or a collection of the axes on the chart.

## Syntax

_expression_.**Axes** (_Type_, _AxisGroup_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Type_ |Optional |**[XlAxisType](excel.xlaxistype.md)** |Specifies the axis to return. The reference style of the formula. Can be one of these **XlAxisType** constants: **xlValue**, **xlCategory**, or **xlSeriesAxis** (valid only for 3D charts).|
|_AxisGroup_ |Optional |**[XlAxisGroup](excel.xlaxisgroup.md)** |The reference style of the formula. Can be one of these **XlAxisGroup** constants: **xlPrimary** or **xlSecondary**. If this argument is omitted, the primary group is used. 3D charts have only one axis group.|

## Example

This example adds an axis label to the category axis.

```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```

<br/>

This example turns off major gridlines for the category axis.

```vb
myChart.Axes(xlCategory).HasMajorGridlines = False
```

<br/>

This example turns off all gridlines for all axes.

```vb
For Each a In myChart.Axes 
 a.HasMajorGridlines = False 
 a.HasMinorGridlines = False 
Next a
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]