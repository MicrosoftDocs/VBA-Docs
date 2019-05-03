---
title: PivotField.DragToColumn property (Excel)
keywords: vbaxl10.chm240102
f1_keywords:
- vbaxl10.chm240102
ms.prod: excel
api_name:
- Excel.PivotField.DragToColumn
ms.assetid: 1e3ce788-5484-2504-37bb-a08770871c98
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.DragToColumn property (Excel)

**True** if the specified field can be dragged to the column position. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**DragToColumn**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

For OLAP data sources, the value is **False** for measure fields.


## Example

This example prevents the Year field in the first PivotTable report on worksheet one from being dragged to the column position.

```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Year").DragToColumn = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]