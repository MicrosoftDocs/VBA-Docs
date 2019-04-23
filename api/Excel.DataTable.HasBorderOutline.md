---
title: DataTable.HasBorderOutline property (Excel)
keywords: vbaxl10.chm626076
f1_keywords:
- vbaxl10.chm626076
ms.prod: excel
api_name:
- Excel.DataTable.HasBorderOutline
ms.assetid: e98c1e9a-ff51-32eb-ab8a-aab92c07c82c
ms.date: 04/23/2019
localization_priority: Normal
---


# DataTable.HasBorderOutline property (Excel)

**True** if the chart data table has outline borders. Read/write **Boolean**.


## Syntax

_expression_.**HasBorderOutline**

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