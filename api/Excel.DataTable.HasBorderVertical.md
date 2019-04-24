---
title: DataTable.HasBorderVertical property (Excel)
keywords: vbaxl10.chm626075
f1_keywords:
- vbaxl10.chm626075
ms.prod: excel
api_name:
- Excel.DataTable.HasBorderVertical
ms.assetid: 166ad9ef-99c1-4e94-079c-4997aacc6e2d
ms.date: 04/23/2019
localization_priority: Normal
---


# DataTable.HasBorderVertical property (Excel)

**True** if the chart data table has vertical cell borders. Read/write **Boolean**.


## Syntax

_expression_.**HasBorderVertical**

_expression_ A variable that represents a **[DataTable](excel.datatable(object).md)** object.


## Example

This example causes the embedded chart data table to be displayed with an outline border and no cell borders.

```vb
With Worksheets(1).ChartObjects(1).Chart 
 .HasDataTable = True 
 With .DataTable 
 .HasBorderHorizontal = False 
 .HasBorderVertical = False 
 .HasBorderOutline = True 
 End With 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]