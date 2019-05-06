---
title: PivotField.ParentField property (Excel)
keywords: vbaxl10.chm240089
f1_keywords:
- vbaxl10.chm240089
ms.prod: excel
api_name:
- Excel.PivotField.ParentField
ms.assetid: 4b609a86-9a25-f292-7446-2a65ea1f90a0
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.ParentField property (Excel)

Returns a **PivotField** object that represents the PivotTable field that's the group parent of the specified object. The field must be grouped and must have a parent field. Read-only.


## Syntax

_expression_.**ParentField**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Example

This example displays the name of the field that's the group parent of the field that contains the active cell.

```vb
Worksheets("Sheet1").Activate 
MsgBox "The active field is a child of the field " & _ 
 ActiveCell.PivotField.ParentField.Name
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]