---
title: DataTable property (Excel Graph)
keywords: vbagr10.chm66931
f1_keywords:
- vbagr10.chm66931
ms.prod: excel
api_name:
- Excel.DataTable
ms.assetid: bf432a3e-dd5e-db5b-63b3-4d037976edcc
ms.date: 04/10/2019
localization_priority: Normal
---


# DataTable property (Excel Graph)

Returns a **DataTable** object that represents the chart data table. Read-only.

## Syntax

_expression_.**DataTable**

_expression_ Required. An expression that returns a **[DataTable](Excel.DataTable-graph-object.md)** object.

## Example

This example adds a data table with an outline border to the chart.

```vb
With myChart 
 .HasDataTable = True 
 .DataTable.HasBorderOutline = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]