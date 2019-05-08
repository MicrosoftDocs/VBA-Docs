---
title: PivotTable.PageFieldOrder property (Excel)
keywords: vbaxl10.chm235119
f1_keywords:
- vbaxl10.chm235119
ms.prod: excel
api_name:
- Excel.PivotTable.PageFieldOrder
ms.assetid: 0c8a6473-f2ee-f357-b840-aaf61cee1fa0
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PageFieldOrder property (Excel)

Returns or sets the order in which page fields are added to the PivotTable report's layout. Can be one of the following **[XlOrder](Excel.XlOrder.md)** constants: **xlDownThenOver** or **xlOverThenDown**. The default constant is **xlDownThenOver**. Read/write **Long**.


## Syntax

_expression_.**PageFieldOrder**

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