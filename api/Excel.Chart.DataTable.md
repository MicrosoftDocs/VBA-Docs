---
title: Chart.DataTable property (Excel)
keywords: vbaxl10.chm149098
f1_keywords:
- vbaxl10.chm149098
ms.prod: excel
api_name:
- Excel.Chart.DataTable
ms.assetid: e977daf1-45a1-a069-3d6c-afbe13724d11
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.DataTable property (Excel)

Returns a  **[DataTable](Excel.DataTable(object).md)** object that represents the chart data table. Read-only.


## Syntax

_expression_. `DataTable`

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Example

This example adds a data table with an outline border to the embedded chart.


```vb
With Worksheets(1).ChartObjects(1).Chart 
 .HasDataTable = True 
 .DataTable.HasBorderOutline = True 
End With
```


## See also


[Chart Object](Excel.Chart(object).md)

