---
title: Chart.HasDataTable property (Excel)
keywords: vbaxl10.chm149114
f1_keywords:
- vbaxl10.chm149114
api_name:
- Excel.Chart.HasDataTable
ms.assetid: c29e7606-086e-8549-2259-332d30c1846a
ms.date: 04/16/2019
ms.localizationpriority: medium
---


# Chart.HasDataTable property (Excel)

**True** if the chart has a data table. Read/write **Boolean**.


## Syntax

_expression_.**HasDataTable**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


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