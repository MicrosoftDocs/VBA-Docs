---
title: DataTable object (Excel Graph)
keywords: vbagr10.chm5207296
f1_keywords:
- vbagr10.chm5207296
ms.prod: excel
api_name:
- Excel.DataTable
ms.assetid: cf9aa637-3b5d-1e18-1956-291a0295dddf
ms.date: 04/06/2019
localization_priority: Normal
---


# DataTable object (Excel Graph)

Represents a data table in the specified chart.


## Remarks

Use the **[DataTable](excel.datatable-graph-property.md)** property to return a **DataTable** object. 


## Example

The following example adds a data table with an outline border to the embedded chart.

```vb
With myChart 
 .HasDataTable = True 
 .DataTable.HasBorderOutline = True 
End With
```


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]