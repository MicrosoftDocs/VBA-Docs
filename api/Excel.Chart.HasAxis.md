---
title: Chart.HasAxis property (Excel)
keywords: vbaxl10.chm149113
f1_keywords:
- vbaxl10.chm149113
ms.prod: excel
api_name:
- Excel.Chart.HasAxis
ms.assetid: f2df9f16-980d-fd02-3e09-6d6903dbb6c6
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.HasAxis property (Excel)

Returns or sets which axes exist on the chart. Read/write **Variant**.


## Syntax

_expression_.**HasAxis** (_Index1_, _Index2_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index1_|Required| **Variant**|The axis type. Series axes apply only to 3D charts. Can be one of the **[XlAxisType](Excel.XlAxisType.md)** constants.|
| _Index2_|Optional| **Variant**|The axis group. 3D charts have only one set of axes. Can be one of the **[XlAxisGroup](Excel.XlAxisGroup.md)** constants.|

## Remarks

You must enter a value for at least one of the parameters when setting this property.

Microsoft Excel may create or delete axes if you change the chart type or the **[Axis.AxisGroup](Excel.Axis.AxisGroup.md)**, **[Chart.AxisGroup](Excel.ChartGroup.AxisGroup.md)**, or **[Series.AxisGroup](Excel.Series.AxisGroup.md)** properties.


## Example

This example turns on the primary value axis for Chart1.

```vb
Charts("Chart1").HasAxis(xlValue, xlPrimary) = True

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]