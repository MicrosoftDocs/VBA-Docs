---
title: PivotTable.PageFieldWrapCount property (Excel)
keywords: vbaxl10.chm235121
f1_keywords:
- vbaxl10.chm235121
ms.prod: excel
api_name:
- Excel.PivotTable.PageFieldWrapCount
ms.assetid: 930bfe25-362e-f907-d593-6898db07f55b
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PageFieldWrapCount property (Excel)

Returns or sets the number of page fields in each column or row in the PivotTable report. Read/write **Long**.


## Syntax

_expression_.**PageFieldWrapCount**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example causes the PivotTable report to draw three page fields in a row before starting a new row.

```vb
With Worksheets(1).PivotTables("Pivot1") 
 .PageFieldOrder = xlOverThenDown 
 .PageFieldWrapCount = 3 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]