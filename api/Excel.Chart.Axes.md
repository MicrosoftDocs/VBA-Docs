---
title: Chart.Axes method (Excel)
keywords: vbaxl10.chm149081
f1_keywords:
- vbaxl10.chm149081
ms.prod: excel
api_name:
- Excel.Chart.Axes
ms.assetid: d0520f61-9aff-894b-9975-37dcb5b5fe3c
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Axes method (Excel)

Returns an object that represents either a single axis or a collection of the axes on the chart.


## Syntax

_expression_.**Axes** (_Type_, _AxisGroup_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **Variant**|Specifies the axis to return. Can be one of the following **[XlAxisType](Excel.XlAxisType.md)** constants: **xlValue**, **xlCategory**, or **xlSeriesAxis** (**xlSeriesAxis** is valid only for 3D charts).|
| _AxisGroup_|Optional| **[XlAxisGroup](Excel.XlAxisGroup.md)**|Specifies the axis group. If this argument is omitted, the primary group is used. 3D charts have only one axis group.|

## Return value

Object


## Example

This example adds an axis label to the category axis on Chart1.

```vb
With Charts("Chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```

<br/>

This example turns off major gridlines for the category axis on Chart1.

```vb
Charts("Chart1").Axes(xlCategory).HasMajorGridlines = False
```

<br/>

This example turns off all gridlines for all axes on Chart1.

```vb
For Each a In Charts("Chart1").Axes 
 a.HasMajorGridlines = False 
 a.HasMinorGridlines = False 
Next a
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
