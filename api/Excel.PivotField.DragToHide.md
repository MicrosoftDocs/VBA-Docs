---
title: PivotField.DragToHide property (Excel)
keywords: vbaxl10.chm240103
f1_keywords:
- vbaxl10.chm240103
ms.prod: excel
api_name:
- Excel.PivotField.DragToHide
ms.assetid: 24bccf39-3271-4387-6b7b-21f0ba47500c
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.DragToHide property (Excel)

**True** if the field can be hidden by being dragged off the PivotTable report. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**DragToHide**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Example

This example prevents the Year field in the first PivotTable report on worksheet one from being dragged off the report.

```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Year").DragToHide = False
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]