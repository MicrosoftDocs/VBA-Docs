---
title: ChartObjects.Add method (Excel)
keywords: vbaxl10.chm497103
f1_keywords:
- vbaxl10.chm497103
ms.prod: excel
api_name:
- Excel.ChartObjects.Add
ms.assetid: 46f28b34-83a5-b3d9-c19b-a1dc8e05dff7
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObjects.Add method (Excel)

Creates a new embedded chart.


## Syntax

_expression_.**Add** (_Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[ChartObjects](Excel.ChartObjects.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Left_|Required| **Double**|The initial coordinates of the new object (in [points](../language/glossary/vbe-glossary.md#point)), relative to the upper-left corner of cell A1 on a worksheet or to the upper-left corner of a chart.|
| _Width_|Required| **Double**|The initial size of the new object, in points.|

## Return value

A **[ChartObject](Excel.ChartObject.md)** object that represents the new embedded chart.


## Example

This example creates a new embedded chart.

```vb
Set co = Sheets("Sheet1").ChartObjects.Add(50, 40, 200, 100) 
co.Chart.ChartWizard Source:=Worksheets("Sheet1").Range("A1:B2"), _ 
 Gallery:=xlColumn, Format:=6, PlotBy:=xlColumns, _ 
 CategoryLabels:=1, SeriesLabels:=0, HasLegend:=1
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
