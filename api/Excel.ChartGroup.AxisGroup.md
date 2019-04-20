---
title: ChartGroup.AxisGroup property (Excel)
keywords: vbaxl10.chm568073
f1_keywords:
- vbaxl10.chm568073
ms.prod: excel
api_name:
- Excel.ChartGroup.AxisGroup
ms.assetid: 2fa4488c-6a50-9aac-affe-a6f2b8afa62e
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.AxisGroup property (Excel)

Returns or sets the group for the specified chart. Read/write.


## Syntax

_expression_.**AxisGroup**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Return value

**[XlAxisGroup](Excel.XlAxisGroup.md)**


## Remarks

For 3D charts, only **xlPrimary** is valid.


## Example

This example deletes the value axis if it is in the secondary group.

```vb
With myChart.Axes(xlValue) 
 If .AxisGroup = xlSecondary Then .Delete 
End With 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]