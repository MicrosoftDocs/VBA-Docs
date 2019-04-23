---
title: HasTitle property (Excel Graph)
keywords: vbagr10.chm65590
f1_keywords:
- vbagr10.chm65590
ms.prod: excel
api_name:
- Excel.HasTitle
ms.assetid: 9ecc48d3-fd86-e185-a69f-0676230b3194
ms.date: 04/11/2019
localization_priority: Normal
---


# HasTitle property (Excel Graph)

**True** if the axis or chart has a visible title. Read/write **Boolean**.

## Syntax

_expression_.**HasTitle**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

An axis title is represented by an **[AxisTitle](Excel.AxisTitle-graph-object.md)** object.

A chart title is represented by a **[ChartTitle](Excel.ChartTitle-graph-object.md)** object.


## Example

This example adds an axis label to the category axis.

```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]